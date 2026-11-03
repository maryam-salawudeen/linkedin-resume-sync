# LinkedIn résumé sync

Reads my résumé (.docx) and builds one page with every LinkedIn field filled in, in the order LinkedIn's edit forms ask for them. Each field has a Copy button. When the résumé changes, the page shows only what moved since I last updated LinkedIn.

It runs on my Mac with plain Python 3 and nothing to install. It never logs in to LinkedIn, never asks for a password and sends nothing anywhere.

## Run it

```bash
cd ~/Developer/linkedin-resume-sync
```

```bash
python3 sync.py
```

That reads `~/Downloads/Maryam_Salawudeen_Resume_Security_2026.docx` and opens `out/linkedin.html`. To use a different file:

```bash
python3 sync.py ~/Downloads/some-other-resume.docx
```

It needs the .docx. The PDF loses the structure the parser relies on.

## Updating LinkedIn

1. Open LinkedIn in the browser where I'm signed in.
2. Work down the page. Each section says where the form is (for example, Profile → Experience → +). Copy a field, paste it, move to the next.
3. Tick Done on a card once I've clicked Save on LinkedIn. The ticks stay if I reload the page.
4. When the profile matches, record it:

```bash
python3 sync.py --done
```

Next time I edit the résumé and run `python3 sync.py`, cards are marked NEW or CHANGED, "Only show changes" is on, and a changed field has a "What it was before" link. Anything I've deleted from the résumé shows in a red box at the top so I can remove it from LinkedIn too.

Skills are chips. Click one to copy it. Right-click to cross it off.

## Letting Claude in Chrome fill the forms

The Copy for Claude in Chrome button copies a prompt with every new or changed field (all of them on a first run). Paste it into Claude in Chrome on a LinkedIn tab. The prompt tells it to fill each form and stop before Save, so I check each one and click Save myself. It also tells it to leave "Share with network" off.

## Things the tool can't fill

- Employment type, education dates and certification issue dates aren't on the résumé. The page says so on each one.
- LinkedIn has no "in progress" option for certifications, so the CISSP card is marked "don't add". It shows in my headline and About instead.
- The Virtual SOC Analyst Prep workshop goes under Courses, which has no date field. The MetaCTF challenge goes under Honors & awards.
- Location, company and school boxes use LinkedIn's own lists, so pick the closest match after pasting.

## The About text

The résumé summary is written in résumé style. My first-person About lives in `overrides.json`, next to the summary it was written from. If I change the summary on the résumé, the page falls back to the new summary text and warns me, so I know to rewrite the About in `overrides.json`.

## Why there's no "sync" button

LinkedIn doesn't let personal tools edit a profile. Its profile-editing API is only open to approved partners, and its user agreement bans bots and automated activity on the site. A script that logs in as me could get the account restricted. So the tool does the slow part (reading and laying out the résumé) and I do the clicking.

## Privacy

- The phone number is removed from everything the tool writes, and a test checks that.
- `out/` (the page) and `state/` (the record of what's on LinkedIn) are in `.gitignore`, so résumé text stays out of git.

## Tests

```bash
python3 -m unittest test_sync
```
