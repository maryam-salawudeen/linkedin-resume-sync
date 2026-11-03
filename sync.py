#!/usr/bin/env python3
"""Turn a .docx résumé into a page of LinkedIn fields, ready to copy.

Usage:
    python3 sync.py [resume.docx]          build out/linkedin.html and open it
    python3 sync.py [resume.docx] --done   record this résumé as "on LinkedIn now"
    python3 sync.py --no-open              build without opening a browser

The page marks every field as new, changed or unchanged compared with the last
time you ran --done, so after a résumé edit you only touch what moved.

Standard library only. Nothing is sent anywhere.
"""

import argparse
import datetime
import html
import json
import os
import re
import sys
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "out"
STATE_FILE = HERE / "state" / "linkedin-snapshot.json"
OVERRIDES_FILE = HERE / "overrides.json"
TEMPLATE_FILE = HERE / "page_template.html"
DEFAULT_RESUME = Path.home() / "Downloads" / "Maryam_Salawudeen_Resume_Security_2026.docx"

# LinkedIn's published character limits for the fields that have one.
LIMITS = {
    "headline": 220,
    "about": 2600,
    "position_title": 100,
    "position_description": 2000,
    "project_description": 2000,
    "education_description": 1000,
}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.? \d{4}"
DATE_RANGE = re.compile(rf"({MONTH})\s*[–—-]\s*({MONTH}|Present|Current)", re.I)
SINGLE_DATE = re.compile(rf"^{MONTH}$", re.I)
PHONE = re.compile(r"(?:\+?1[\s.-]?)?(?:\(\d{3}\)|\b\d{3})[\s.-]?\d{3}[\s.-]\d{4}\b")


# ---------------------------------------------------------------- reading the docx

def read_paragraphs(docx_path):
    """Return [(text, is_heading, is_list_item)] for each non-empty paragraph."""
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    paragraphs = []
    for p in root.iter(W + "p"):
        parts = []
        for node in p.iter():
            if node.tag == W + "t":
                parts.append(node.text or "")
            elif node.tag == W + "tab":
                parts.append(" ")
            elif node.tag == W + "br":
                parts.append("\n")
        text = re.sub(r"[ \t]+", " ", "".join(parts)).strip()
        if not text:
            continue
        style = p.find(f"{W}pPr/{W}pStyle")
        style = style.get(W + "val") if style is not None else ""
        is_list = p.find(f"{W}pPr/{W}numPr") is not None or style.startswith("List")
        is_heading = style.lower().startswith("heading")
        paragraphs.append((text, is_heading, is_list))
    return paragraphs


def split_sections(paragraphs):
    """Group paragraphs under their headings. Lines before the first heading go to HEADER."""
    sections = {"HEADER": []}
    current = "HEADER"
    for i, (text, is_heading, is_list) in enumerate(paragraphs):
        # Line 0 is the name, usually in capitals too, so it never counts as a heading.
        looks_like_heading = i > 0 and not is_list and text.isupper() and len(text) < 40
        if is_heading or looks_like_heading:
            current = text.upper()
            sections.setdefault(current, [])
        else:
            sections[current].append((text, is_list))
    return sections


def find_section(sections, *names):
    for name in names:
        for key, value in sections.items():
            if key.startswith(name):
                return value
    return []


# ---------------------------------------------------------------- parsing each part

def parse_header(lines):
    name = lines[0][0].title() if lines else ""
    headline = lines[1][0] if len(lines) > 1 else ""
    city, website = "", ""
    for text, _ in lines[2:]:
        parts = [p.strip() for p in text.split("|")]
        for part in parts:
            if "github.com" in part or (re.search(r"\.\w{2,}(/|$)", part) and "linkedin.com" not in part and "@" not in part):
                website = website or part
        # The location is the first piece of the contact line that isn't a phone, email or link.
        if not city and "@" in text:
            for part in parts:
                if not PHONE.search(part) and "@" not in part and "." not in part:
                    city = part
                    break
    return {"name": name, "headline": headline, "city": city, "website": website}


def parse_experience(lines):
    jobs, job = [], None
    for text, is_list in lines:
        if is_list and job:
            job["bullets"].append(text)
        elif DATE_RANGE.search(text) and job and not job["company"]:
            where, _, when = text.rpartition("|") if "|" in text else ("", "", text)
            company, _, location = where.strip().partition(",")
            start, end = DATE_RANGE.search(when).groups()
            job.update(company=company.strip(), location=location.strip(), start=start, end=end)
        else:
            job = {"title": text, "company": "", "location": "", "start": "", "end": "", "bullets": []}
            jobs.append(job)
    return jobs


def parse_education(lines):
    schools = []
    for text, _ in lines:
        award, _, where = text.partition("|")
        school, _, location = where.strip().partition(",")
        degree, sep, field = award.strip().partition(" in ")
        schools.append({
            "school": school.strip() or award.strip(),
            "location": location.strip(),
            "degree": degree.strip(),
            "field": field.strip() if sep else "",
        })
    return schools


def parse_certifications(lines):
    certs = []
    for text, _ in lines:
        in_progress = bool(re.search(r"\bin progress\b", text, re.I))
        clean = re.sub(r"[:,(]?\s*in progress\)?\s*$", "", text, flags=re.I).strip()
        name, sep, issuer = clean.rpartition(", ")
        if not sep or "(" in issuer:
            name, issuer = clean, ""
        certs.append({"name": name, "issuer": issuer, "in_progress": in_progress})
    return certs


def parse_training(lines):
    items = []
    for text, _ in lines:
        parts = [p.strip() for p in text.split(", ")]
        date = parts.pop() if parts and SINGLE_DATE.match(parts[-1]) else ""
        name = parts.pop(0) if parts else text
        is_competition = re.search(r"\b(CTF|competition|competitor|challenge|hackathon)\b", text, re.I)
        items.append({
            "name": name,
            "details": parts,          # provider, role, etc. in résumé order
            "date": date,
            "kind": "competition" if is_competition else "course",
        })
    return items


def parse_projects(lines):
    projects, project = [], None
    for text, is_list in lines:
        if is_list and project:
            project["bullets"].append(text)
        elif project and not project["bullets"] and not project["summary"] and "|" not in text:
            project["summary"] = text
        else:
            name, _, stack = text.partition("|")
            pieces = [re.sub(r"^built with\s+", "", s.strip(), flags=re.I) for s in split_list(stack)]
            project = {
                "name": name.strip(),
                "stack": pieces,
                "skills": [s for s in pieces if s.lower() != "claude code"],
                "summary": "",
                "bullets": [],
            }
            projects.append(project)
    return projects


def parse_skills(lines):
    groups = []
    for text, _ in lines:
        label, sep, rest = text.partition(":")
        if not sep:
            label, rest = "Skills", text
        # LinkedIn lists skills in sentence case, so "patch management" becomes "Patch management".
        groups.append({"group": label.strip(), "skills": [s[0].upper() + s[1:] for s in split_list(rest)]})
    return groups


def split_list(text):
    """Split on commas that aren't inside parentheses."""
    items, depth, buf = [], 0, ""
    for ch in text:
        depth += ch == "("
        depth -= ch == ")"
        if ch == "," and depth == 0:
            items.append(buf.strip())
            buf = ""
        else:
            buf += ch
    items.append(buf.strip())
    return [i for i in items if i]


def parse_resume(docx_path):
    sections = split_sections(read_paragraphs(docx_path))
    known = ("HEADER", "SUMMARY", "ABOUT", "PROFILE", "SKILLS", "CERTIFICATION", "TRAINING",
             "EXPERIENCE", "WORK", "PROJECT", "PERSONAL PROJECT", "EDUCATION")
    return {
        "header": parse_header(sections["HEADER"]),
        "summary": " ".join(t for t, _ in find_section(sections, "SUMMARY", "ABOUT", "PROFILE")),
        "experience": parse_experience(find_section(sections, "EXPERIENCE", "WORK")),
        "education": parse_education(find_section(sections, "EDUCATION")),
        "certifications": parse_certifications(find_section(sections, "CERTIFICATION", "LICENSE")),
        "training": parse_training(find_section(sections, "TRAINING", "COURSES")),
        "projects": parse_projects(find_section(sections, "PERSONAL PROJECT", "PROJECT")),
        "skills": parse_skills(find_section(sections, "SKILLS")),
        "unrecognised": {k: [t for t, _ in v] for k, v in sections.items()
                         if v and not any(k.startswith(n) for n in known)},
    }


# ---------------------------------------------------------------- LinkedIn fields

def field(label, value, limit=None, note=""):
    return {"label": label, "value": value, "limit": limit, "note": note}


def bullets(lines, prefix=""):
    body = "\n".join(f"• {line}" for line in lines)
    return f"{prefix}\n\n{body}".strip() if prefix else body


def project_intro(project):
    lines = [project["summary"].rstrip(".") + "." if project["summary"] else ""]
    if project["stack"]:
        lines.append("Stack: " + ", ".join(project["stack"]) + ".")
    return "\n".join(line for line in lines if line)


def build_linkedin(resume, overrides):
    """Lay the résumé out as LinkedIn's edit forms ask for it, in their order."""
    sections = []
    h = resume["header"]

    sections.append({
        "id": "intro", "title": "Intro", "where": "Profile → pencil icon next to your name",
        "entries": [{"key": "intro", "title": h["name"] or "Intro", "fields": [
            field("Headline", h["headline"], LIMITS["headline"]),
            field("City", h["city"], note="Pick the closest match from LinkedIn's list." if h["city"]
                  else "No city on your résumé, so leave LinkedIn's location as it is."),
            field("Website (Contact info)", h["website"], note="Under Contact info → Website. Type: Portfolio or Other."),
        ]}],
    })

    about = resume["summary"]
    about_note = "Copied from your résumé summary."
    override = overrides.get("about") or {}
    if override.get("text"):
        if override.get("based_on_summary", "").strip() == about.strip():
            about, about_note = override["text"], "First-person version from overrides.json."
        else:
            about_note = ("Your résumé summary changed since the first-person About in overrides.json was "
                          "written, so this is the résumé text. Rewrite overrides.json to match.")
    sections.append({
        "id": "about", "title": "About", "where": "Profile → About → pencil icon",
        "entries": [{"key": "about", "title": "About", "fields": [
            field("About", about, LIMITS["about"], about_note),
        ]}],
    })

    sections.append({
        "id": "experience", "title": "Experience", "where": "Profile → Experience → + (or pencil on the role)",
        "entries": [{
            "key": f"{job['company']}|{job['title']}", "title": f"{job['title']}, {job['company']}",
            "fields": [
                field("Title", job["title"], LIMITS["position_title"]),
                field("Employment type", "", note="Not on your résumé. Pick it from the list."),
                field("Company or organization", job["company"]),
                field("Location", job["location"]),
                field("Start date", job["start"], note="Two dropdowns: month, then year."),
                field("End date", "" if job["end"].lower() in ("present", "current") else job["end"],
                      note="Tick “I am currently working in this role”." if job["end"].lower() in ("present", "current") else "Two dropdowns: month, then year."),
                field("Description", bullets(job["bullets"]), LIMITS["position_description"]),
            ],
        } for job in resume["experience"]],
    })

    sections.append({
        "id": "education", "title": "Education", "where": "Profile → Education → +",
        "entries": [{
            "key": school["school"], "title": school["school"],
            "fields": [
                field("School", school["school"]),
                field("Degree", school["degree"]),
                field("Field of study", school["field"]),
                field("Start and end dates", "", note="Not on your résumé. Leave blank or fill from memory."),
            ],
        } for school in resume["education"]],
    })

    cert_entries = []
    for cert in resume["certifications"]:
        note = "Issue date and credential ID aren't on your résumé. Add them from your ISC2 or INE account if you like."
        if cert["in_progress"]:
            note = ("In progress. LinkedIn has no “in progress” option here, so don't add it as a certification. "
                    "It already shows in your headline and About.")
        elif not cert["issuer"]:
            note = "Your résumé names no separate issuer. Pick the organisation from LinkedIn's list. " + note
        cert_entries.append({"key": cert["name"], "title": cert["name"], "skip": cert["in_progress"], "fields": [
            field("Name", cert["name"]),
            field("Issuing organization", cert["issuer"]),
            field("Issue date", "", note=note),
        ]})
    sections.append({
        "id": "certifications", "title": "Licenses & certifications",
        "where": "Profile → Licenses & certifications → +", "entries": cert_entries,
    })

    courses = [t for t in resume["training"] if t["kind"] == "course"]
    if courses:
        sections.append({
            "id": "courses", "title": "Courses", "where": "Profile → Add profile section → Recommended → Add courses",
            "entries": [{"key": t["name"], "title": t["name"], "fields": [
                field("Course name", ", ".join([t["name"]] + t["details"]),
                      note=f"LinkedIn's course form has no date field. Your résumé date: {t['date']}." if t["date"] else ""),
            ]} for t in courses],
        })

    competitions = [t for t in resume["training"] if t["kind"] == "competition"]
    if competitions:
        sections.append({
            "id": "honors", "title": "Honors & awards",
            "where": "Profile → Add profile section → Additional → Add honors & awards",
            "entries": [{"key": t["name"], "title": t["name"], "fields": [
                field("Title", t["name"]),
                field("Issuer", t["details"][0] if t["details"] else ""),
                field("Issue date", t["date"], note="Two dropdowns: month, then year."),
                field("Description", ". ".join(d[:1].upper() + d[1:] for d in t["details"][1:]) + ("." if t["details"][1:] else "")),
            ]} for t in competitions],
        })

    sections.append({
        "id": "projects", "title": "Projects", "where": "Profile → Add profile section → Recommended → Add projects",
        "entries": [{
            "key": p["name"], "title": p["name"],
            "fields": [
                field("Project name", p["name"]),
                field("Description", bullets(p["bullets"], prefix=project_intro(p)), LIMITS["project_description"]),
                field("Skills", ", ".join(p["skills"]), note="Add the ones LinkedIn's list offers."),
            ],
        } for p in resume["projects"]],
    })

    sections.append({
        "id": "skills", "title": "Skills", "where": "Profile → Skills → + (add one at a time)", "kind": "chips",
        "entries": [{"key": s, "title": s, "group": g["group"], "fields": [field("Skill", s)]}
                    for g in resume["skills"] for s in g["skills"]],
    })

    if resume["unrecognised"]:
        sections.append({
            "id": "other", "title": "Résumé sections this tool doesn't map", "where": "Check these by hand",
            "entries": [{"key": k, "title": k.title(), "fields": [field("Text", "\n".join(v))]}
                        for k, v in resume["unrecognised"].items()],
        })
    return sections


# ---------------------------------------------------------------- comparing with last time

def snapshot_of(sections):
    return {s["id"]: {e["key"]: {f["label"]: f["value"] for f in e["fields"]} for e in s["entries"]}
            for s in sections}


def mark_changes(sections, previous):
    """Tag every entry and field as new/changed/same, and list what's gone."""
    removed = []
    for section in sections:
        before = (previous or {}).get(section["id"], {})
        for entry in section["entries"]:
            old = before.get(entry["key"])
            if previous is None:
                entry["status"] = "new"
            elif old is None:
                entry["status"] = "new"
            else:
                entry["status"] = "same"
            for f in entry["fields"]:
                if old is not None and f["label"] in old and old[f["label"]] != f["value"]:
                    f["status"], f["old"] = "changed", old[f["label"]]
                    entry["status"] = "changed"
                else:
                    f["status"] = "new" if entry["status"] == "new" else "same"
        keys = {e["key"] for e in section["entries"]}
        for key in before:
            if key not in keys:
                removed.append({"section": section["title"], "title": key.replace("|", " · ")})
    return removed


# ---------------------------------------------------------------- output

def scrub(text):
    return PHONE.sub("[phone removed]", text)


def render(sections, removed, meta):
    data = json.dumps({"sections": sections, "removed": removed, "meta": meta}, ensure_ascii=False)
    data = scrub(data).replace("</", "<\\/")
    page = TEMPLATE_FILE.read_text(encoding="utf-8")
    return page.replace("/*__DATA__*/null", data).replace("__TITLE__", html.escape(meta["name"] or "LinkedIn fields"))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("resume", nargs="?", default=os.environ.get("RESUME", str(DEFAULT_RESUME)),
                    help="path to the .docx résumé (default: %(default)s)")
    ap.add_argument("--done", action="store_true",
                    help="record this résumé as what's on LinkedIn now, so the next run only shows later changes")
    ap.add_argument("--no-open", action="store_true", help="don't open the page in a browser")
    args = ap.parse_args(argv)

    resume_path = Path(args.resume).expanduser()
    if resume_path.suffix.lower() != ".docx":
        sys.exit(f"Give me the .docx version of the résumé (got {resume_path.name}). PDFs don't keep the structure.")
    if not resume_path.exists():
        sys.exit(f"Can't find {resume_path}")

    overrides = json.loads(OVERRIDES_FILE.read_text(encoding="utf-8")) if OVERRIDES_FILE.exists() else {}
    resume = parse_resume(resume_path)
    sections = build_linkedin(resume, overrides)

    state = json.loads(STATE_FILE.read_text(encoding="utf-8")) if STATE_FILE.exists() else None
    if args.done:
        STATE_FILE.parent.mkdir(exist_ok=True)
        state = {"saved": datetime.date.today().isoformat(), "resume": resume_path.name,
                 "fields": json.loads(scrub(json.dumps(snapshot_of(sections), ensure_ascii=False)))}
        STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Recorded. Next run compares against today ({state['saved']}).")

    removed = mark_changes(sections, state["fields"] if state else None)
    changed = sum(e["status"] != "same" for s in sections for e in s["entries"] if not e.get("skip"))
    meta = {
        "name": resume["header"]["name"],
        "resume": resume_path.name,
        "built": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "since": state["saved"] if state else None,
        "changed": changed,
    }

    OUT_DIR.mkdir(exist_ok=True)
    page = OUT_DIR / "linkedin.html"
    page.write_text(render(sections, removed, meta), encoding="utf-8")
    (OUT_DIR / "fields.json").write_text(scrub(json.dumps(sections, indent=2, ensure_ascii=False)), encoding="utf-8")

    if state:
        print(f"{changed} item(s) new or changed and {len(removed)} removed since {state['saved']}.")
    else:
        print("First run: everything is marked new. After you've updated LinkedIn, run with --done.")
    print(f"Page: {page}")
    if not args.no_open:
        webbrowser.open(page.as_uri())


if __name__ == "__main__":
    main()
