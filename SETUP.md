# Setting up your Spending Tracker app

Two parts to deploy: the backend (Apps Script, talks to your Sheets) and the frontend (the webpage, hosted on GitHub Pages). Do the backend first — the frontend needs its URL.

## Part 1 — Backend (Google Apps Script)

1. Go to **script.google.com** and click **New project**.
2. Delete the placeholder code, then paste in the entire contents of `backend/Code.gs`.
3. Rename the project (top left, "Untitled project") to something like "Spending App Backend".
4. In the function dropdown near the Run button, select **setup**, then click **Run**.
   - The first time, Google will ask you to authorize the script — click through (it'll warn it's "unverified," that's normal for your own personal script; click **Advanced > Go to Spending App Backend (unsafe)** and allow it).
   - Check the **Execution log** (View > Logs) for "Setup complete." If it instead says it couldn't find your template, make sure a file named exactly **CURRENT Monthly Spending Template** exists in your Drive, or edit the `TEMPLATE_NAME` constant at the top of Code.gs to match your actual file name.
5. Click **Deploy > New deployment**.
   - Click the gear icon next to "Select type" and choose **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, authorize again if asked.
6. Copy the **Web app URL** it gives you (looks like `https://script.google.com/macros/s/.../exec`). You'll need it in Part 2.

Whenever you edit `Code.gs` later, you'll need to **Deploy > Manage deployments > edit (pencil) > New version** to push the changes live — just saving the file isn't enough.

## Part 2 — Frontend (GitHub Pages)

1. Open `frontend/index.html` in a text editor and replace `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` with the URL you copied in step 6 above. Save it.
2. Go to **github.com**, click the **+** in the top right → **New repository**. Name it something like `spending-app`, keep it **Private**, click **Create repository**.
3. On the new repo's page, click **uploading an existing file** (or drag-and-drop). Drag in everything **inside** the `frontend` folder (`index.html`, `manifest.json`, `sw.js`, and the `icons` folder) — not the `frontend` folder itself, its contents. Commit the upload.
4. Go to the repo's **Settings > Pages**.
   - Under "Build and deployment", Source: **Deploy from a branch**.
   - Branch: **main**, folder: **/(root)**. Save.
5. Wait a minute, then refresh — GitHub shows the live URL at the top (something like `https://yourusername.github.io/spending-app/`).
6. Open that URL on your phone (Safari on iPhone, Chrome on Android), then:
   - **iPhone:** tap the Share icon → **Add to Home Screen**.
   - **Android (Chrome):** tap the ⋮ menu → **Add to Home screen** / **Install app**.

You now have an app icon that opens the logger full-screen.

## Testing it

Log one throwaway entry first (e.g. $0.01, "test") and check it landed correctly in the right spreadsheet tab before you rely on it for real entries — the category/account detection reads your Summary tab's layout, and it's worth confirming it found the right columns before you get several days deep into logging. If something's off, tell me what you're seeing and I'll adjust `Code.gs`.

## Updating later

- **Frontend change:** edit the file, re-upload to GitHub (or use `git push` if you're comfortable with it), Pages updates automatically in about a minute.
- **Backend change:** edit `Code.gs` in the Apps Script editor, then Deploy > Manage deployments > New version (see above) — this part doesn't auto-update.
