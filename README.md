# alpharetta scouting

Scouting site for Alpharetta High School policy debate. It shows what every team on the circuit reads (affs, off-case, block and 2NR choices, new arguments), lets us search card tags and cites across every speech doc, and collects our own RFD notes from a Google Form.

Built by Akshan for the Alpharetta debate team.

## How it works

```
email chains (team docs Gmail)
   │  Apps Script, every 10 min
   ▼
Drive: every .docx saved + its outline (pockets / hats / blocks / tags + cites)
weekly feed sheets "Chain Index <Monday>"
   │  tagging, Mon + Thu: speech, positions, new arguments
   ▼
"Alpharetta Scouting Data" sheet
   │  Apps Script web app (doGet) serves JSON
   ▼
this site (GitHub Pages)  ◄──  RFD Google Form answers (live)
```

- `index.html`: the whole site (HTML, CSS and JS in one file, no build step).
- `config.js`: the only settings file (web app URL, form link).
- `data/`: snapshot of everything indexed before Oct 5, 2026.
- `apps-script/`: the Google Apps Script that runs the pipeline.

## Setup

1. **Script.** In script.google.com (signed in to the team docs account) paste `apps-script/ScoutingIndexer.gs`, run `setUp` once, then Deploy → New deployment → Web app → Execute as *Me*, access *Anyone*. Put the `/exec` URL in `config.js` as `apiUrl`.
2. **Site.** Upload this folder to a GitHub repo, then Settings → Pages → Deploy from branch → `main` / root.

After changing the script later: Deploy → Manage deployments → edit → New version.

Font: Geist and Geist Mono (SIL Open Font License, see `fonts/OFL.txt`).

## Keyboard

- `/` jumps to search, `Esc` clears it.
