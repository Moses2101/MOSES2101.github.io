# Academic website, built from my CV

My Word CV is the only thing I edit. Everything on the website (Home, About,
Research, Teaching, CV and Contact) is generated from it, and a downloadable
copy of the CV is published alongside the pages.

## The update loop

1. **Edit my CV in Word** as usual and save it in my CV folder
   (`...\JESUS jobs application\Moses CV`). The newest `.docx` in that folder
   is the one that gets published, so I can keep dated versions there.
2. **Run one command** from this folder:

   ```
   python publish.py
   ```

   It copies the newest CV in, lists which paragraphs changed, rebuilds the
   site, asks me to confirm, then commits and pushes.
3. **Wait a minute or two.** GitHub rebuilds and redeploys the site on its own.
   Progress shows under the repository's **Actions** tab.

Useful options:

| Command | What it does |
| --- | --- |
| `python publish.py --remove-phone` | Removes any phone number from the published copy of the CV (my original file is untouched). |
| `python publish.py "D:\Another folder"` | Takes the newest CV from a different folder. |
| `python publish.py --no-push` | Commits locally, so I can look first and run `git push` later. |
| `python publish.py --yes` | Skips the "Publish this version?" question. |

## Looking at the site before publishing

```
python serve.py
```

This rebuilds the site and opens it at http://localhost:8000. Addresses work
the same way as on GitHub Pages (`/about`, `/research`, ...). Press Ctrl+C to
stop it.

## Photos

My original photos live in `Images/`. That folder is never published or
committed. The site uses small copies in `assets/photos/`.

To add, swap or remove a photo:

1. Put the original in `Images/`.
2. Open `photos.json` and add an entry where I want it: `portrait` (Home page
   and Contact), `home` (Highlights), `research-presentations`, `research-epc`
   or `about-service`. Each entry has the file name (`src`), a `caption` shown
   under the photo, and `alt` text describing the photo for screen readers.
3. Make the web copies, then preview and publish:

   ```
   powershell -ExecutionPolicy Bypass -File optimize-photos.ps1
   python serve.py
   git add -A
   git commit -m "Update photos"
   git push
   ```

The copies are resized, turned upright, and saved without the original's
hidden details (location, camera, date).

## Rules the build follows

- Phone numbers and home addresses never appear on the pages. If something
  that looks like a phone number reaches a page, the build stops instead of
  publishing it.
- Every section heading in the CV must be a paragraph on its own line. The
  headings it recognizes are listed in `SECTION_HEADINGS` at the top of
  `build.py`. If I rename a heading in Word, I add the new wording there.
- These sections are required, and the build stops with a clear message if one
  is missing: Professional Summary, Education, Research Experience, Teaching
  Experience and Publications.
- Within a section, an entry's date goes at the end of its first line, after
  a tab (for example `Research Assistant<Tab>2023-Present`). Publications are
  numbered lists under their own subheadings (such as "Book chapter").

## One-time setup

1. Install **Python 3** (python.org, or `winget install Python.Python.3.12`)
   and **Git** (`winget install Git.Git`).
2. Create an empty repository on GitHub. Naming it
   `<my-username>.github.io` gives the site the address
   `https://<my-username>.github.io/`. Any other name gives
   `https://<my-username>.github.io/<name>/`.
3. In this folder:

   ```
   git init -b main
   git add -A
   git commit -m "First version of the site"
   git remote add origin https://github.com/<my-username>/<repository>.git
   git push -u origin main
   ```

4. On GitHub, open the repository's **Settings > Pages** and set **Source**
   to **GitHub Actions**. The first deploy then runs on its own, and every
   later push redeploys.

The repository is public, so the CV committed in `cv/cv.docx` is public too.
Use `--remove-phone` if the CV ever includes a phone number.

## What's in this folder

| Path | Purpose |
| --- | --- |
| `cv/cv.docx` | The published copy of my CV (replaced by `publish.py`). |
| `build.py` | Reads the CV and writes the site into `dist/`. |
| `serve.py` | Local preview with clean addresses. |
| `publish.py` | The one-command update. |
| `assets/style.css`, `assets/site.js` | The site's stylesheet and its small script (theme switch, mobile menu, publication filter). |
| `.github/workflows/deploy.yml` | Builds and deploys to GitHub Pages on every push to `main`. |
