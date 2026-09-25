# jennywu-portfolio

Portfolio website for Jenny Wu: a static HTML site with no build step.

## Viewing locally

Open `index.html` in a browser, or serve the folder:

```
python -m http.server 8000
```

then visit http://localhost:8000.

## Structure

```
index.html           Home page
case-*.html          Case studies (AI, DOME, MuseKey, TopJob)
linktree.html        Link-in-bio page
profile.html         Profile page
assets/              Images and videos used by the site, one folder per case study
  home/ ai/ dome/ musekey/ topjob/
cv/                  CV files, VIVAIA application material and build_cv.py
  drafts/            Earlier CV drafts (EN/CN)
archive/             Files not used by the site (source PNGs, raw videos, hero candidates, WorkBuddy notes)
```
