# Tie-Out — Setup Guide

This is the no-build version. No npm. No webpack. Three real files
that matter: `taskpane.html`, `manifest.xml`, and the `assets/` icons.

You will:
1. Put these files in a GitHub repo.
2. Turn on GitHub Pages so the files have a public HTTPS URL.
3. Edit `manifest.xml` once to point at your URL.
4. Sideload the manifest into Word.

After that, "updating the add-in" = edit `taskpane.html`, push to GitHub.
Word picks up the new version next time the pane opens.

---

## Part 1 — Get the files into a GitHub repo

### A. Make a GitHub account (skip if you have one)

Go to <https://github.com> → **Sign up**. The free plan is fine.

### B. Create the repo

1. Top-right **+** menu → **New repository**.
2. Name it whatever — e.g. `tieout`. Lowercase, no spaces.
3. **Public** (required for the free GitHub Pages tier; private works
   on paid plans).
4. **Add a README** — checked.
5. Click **Create repository**.

### C. Upload the files (browser, no command line)

1. In the new repo, click **Add file → Upload files**.
2. Drag in `taskpane.html`, `manifest.xml`, and the entire `assets`
   folder (all three icon PNGs).
3. Scroll down → **Commit changes**.

You should see this layout in the repo:
```
tieout/
├─ README.md
├─ taskpane.html
├─ manifest.xml
└─ assets/
   ├─ icon-16.png
   ├─ icon-32.png
   └─ icon-80.png
```

---

## Part 2 — Turn on GitHub Pages

This is what gives your files a public HTTPS URL that Word can reach.

1. In the repo, click **Settings** (top tab).
2. Left sidebar → **Pages**.
3. **Source**: select **Deploy from a branch**.
4. **Branch**: select **main**, folder **/(root)**. Click **Save**.
5. Wait ~1 minute. Refresh. Near the top you'll see:

   *Your site is live at* `https://YOUR-USERNAME.github.io/YOUR-REPO/`

   Copy that URL.

### Verify

Paste `https://YOUR-USERNAME.github.io/YOUR-REPO/taskpane.html` into a
browser. You should see the Tie-Out pane (it'll show "This add-in must
be run inside Word." in the status bar — that's expected, it just means
Office.js can tell it isn't actually inside Word).

If you see a 404, give it another minute — first deploys can take a few.

---

## Part 3 — Edit the manifest

Open `manifest.xml` in the repo (click the file, then the pencil icon
to edit in the browser). Use **Find & replace** (Ctrl/Cmd-F) to swap:

- `YOUR-USERNAME` → your GitHub username
- `YOUR-REPO` → your repo name

There are about 8 occurrences. Commit changes.

Now download the edited `manifest.xml` to your computer — click the
file, then the **Download raw file** button. Save it somewhere
findable, like Desktop.

---

## Part 4 — Sideload into Word

You only do this once per machine. The manifest is a pointer; Word
will fetch the actual code from GitHub Pages every time the pane
opens.

### Easy way (M365 admin) — for tenant-wide deployment

If you have admin rights at Baylor Genetics:
1. <https://admin.microsoft.com> → **Settings → Integrated apps**.
2. **Upload custom apps** → upload `manifest.xml`.
3. Assign to yourself (and anyone else who needs it).
4. Wait ~24 hours for it to appear; it will show up in Word's ribbon
   automatically.

This is the right answer for a real rollout. Until then:

### Sideload (just you, your machine) — Windows

1. Create a folder, e.g. `C:\tieout-manifest`.
2. Drop `manifest.xml` into it.
3. Right-click that folder → **Properties** → **Sharing** tab → **Share**
   → add yourself → **Share**. Note the network path
   (`\\YOUR-PC\tieout-manifest`).
4. Open Word → **File → Options → Trust Center → Trust Center Settings
   → Trusted Add-in Catalogs**.
5. Paste the share path into **Catalog URL**, click **Add catalog**,
   tick **Show in Menu**, click **OK**.
6. Restart Word.
7. **Insert → My Add-ins → SHARED FOLDER** tab. You'll see "Tie-Out".
   Click → **Add**.

### Sideload — Mac

1. Open Finder, press **Cmd-Shift-G**, paste:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   (Create the `wef` folder if missing.)
2. Copy `manifest.xml` into that folder.
3. Restart Word.
4. **Insert → My Add-ins → Developer Add-ins**. You'll see "Tie-Out".

---

## Using it

1. Open any Word doc.
2. Click the **Tie-Out** ribbon tab → **Open Pane**.
3. In the pane: **Load PDF…**, pick a PDF.
4. Click somewhere on the PDF page — an oxblood circle marks your
   anchor point.
5. In Word, select the text you want to tie to that PDF location.
6. Click **Create tie-out from selection**. The Word selection gets
   wrapped in a red bounding box and shows up as an exhibit card in
   the right column.
7. Click any exhibit card to jump back to that text + that PDF page.
   The `↗ open` button launches the PDF in your default external app
   (Adobe) at the correct page.

---

## Updating the add-in later

Just edit `taskpane.html` in the GitHub repo and commit. GitHub Pages
re-deploys automatically in a minute or two. Next time you open the
Tie-Out pane in Word, it'll fetch the new version.

You only have to re-sideload `manifest.xml` if you change something in
the manifest itself (ribbon labels, tab name, permissions). Day-to-day
feature work doesn't require it.

---

## Caching gotcha

Word and Office both aggressively cache add-in code. If you push an
update and Word still shows the old version:

- Close the pane and reopen it.
- If that doesn't work: in Word → **Insert → My Add-ins → ⋯ menu next
  to Tie-Out → Refresh**.
- Nuclear option: clear the Office cache.
  - Windows: delete contents of
    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
  - Mac: delete contents of
    `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`
  - Then restart Word.

---

## Known limitations (Phase 1)

- The browser File API doesn't expose real file paths. The `↗ open`
  button uses just the file name, so it only works if Adobe finds the
  PDF in its default open path. Phase 2 will add proper SharePoint URL
  support.
- PDF only. Excel/image/other-Word-doc anchors are Phase 2+.
- No image hover preview in Word — by design. The pane plays that role.
