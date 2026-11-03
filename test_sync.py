"""Checks the parser and change tracking on a small made-up résumé.

Run: python3 -m unittest test_sync
"""

import tempfile
import unittest
import zipfile
from pathlib import Path

import sync

DOC = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}</w:body></w:document>"""


def para(text, style=None, listed=False):
    ppr = ""
    if style or listed:
        ppr = "<w:pPr>" + (f'<w:pStyle w:val="{style}"/>' if style else "") + \
              ('<w:numPr><w:numId w:val="1"/></w:numPr>' if listed else "") + "</w:pPr>"
    return f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def make_docx(path, title="Analyst", extra_skill=""):
    body = "".join([
        para("JANE DOE"),
        para("Security Analyst | SSCP"),
        para("Toronto, ON | 416-555-0199 | jane@example.com"),
        para("SUMMARY", "Heading1"),
        para("Analyst with 3 years in IAM."),
        para("SKILLS", "Heading1"),
        para("Identity: Active Directory, single sign-on (SSO, SAML)" + extra_skill),
        para("CERTIFICATIONS", "Heading1"),
        para("Systems Security Certified Practitioner (SSCP), ISC2", listed=True),
        para("CISSP, ISC2: in progress", listed=True),
        para("EXPERIENCE", "Heading1"),
        para(title),
        para("Acme Corp., Toronto, ON | Jan 2021 – Present"),
        para("Ran access reviews.", listed=True),
        para("EDUCATION", "Heading1"),
        para("Bachelor of Science in Computer Science | Example University, Toronto, ON"),
    ])
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", DOC.format(body))


class SyncTest(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.docx = self.dir / "r.docx"

    def build(self, **kw):
        make_docx(self.docx, **kw)
        return sync.build_linkedin(sync.parse_resume(self.docx), {})

    def test_parses_sections(self):
        make_docx(self.docx)
        r = sync.parse_resume(self.docx)
        self.assertEqual(r["header"]["name"], "Jane Doe")
        self.assertEqual(r["header"]["city"], "Toronto, ON")
        job = r["experience"][0]
        self.assertEqual((job["title"], job["company"], job["start"], job["end"]),
                         ("Analyst", "Acme Corp.", "Jan 2021", "Present"))
        self.assertEqual(r["skills"][0]["skills"], ["Active Directory", "Single sign-on (SSO, SAML)"])
        self.assertTrue(r["certifications"][1]["in_progress"])
        self.assertEqual(r["certifications"][1]["name"], "CISSP")
        self.assertEqual(r["education"][0]["field"], "Computer Science")

    def test_phone_never_reaches_output(self):
        sections = self.build()
        page = sync.render(sections, [], {"name": "x", "resume": "r", "built": "", "since": None, "changed": 0})
        self.assertNotIn("555-0199", page)
        self.assertNotIn("555-0199", sync.scrub("call 416-555-0199"))

    def test_marks_changes_since_last_done(self):
        before = sync.snapshot_of(self.build())
        after = self.build(title="Senior Analyst", extra_skill=", Okta")
        removed = sync.mark_changes(after, before)
        status = {(s["id"], e["key"]): e["status"] for s in after for e in s["entries"]}
        self.assertEqual(status[("skills", "Okta")], "new")
        self.assertEqual(status[("skills", "Active Directory")], "same")
        self.assertEqual(status[("experience", "Acme Corp.|Senior Analyst")], "new")
        self.assertIn({"section": "Experience", "title": "Acme Corp. · Analyst"}, removed)

    def test_field_edit_is_changed_not_new(self):
        before = sync.snapshot_of(self.build())
        make_docx(self.docx)
        r = sync.parse_resume(self.docx)
        r["experience"][0]["bullets"].append("Wrote access policy.")
        after = sync.build_linkedin(r, {})
        sync.mark_changes(after, before)
        job = next(s for s in after if s["id"] == "experience")["entries"][0]
        self.assertEqual(job["status"], "changed")
        desc = next(f for f in job["fields"] if f["label"] == "Description")
        self.assertEqual(desc["old"], "• Ran access reviews.")


if __name__ == "__main__":
    unittest.main()
