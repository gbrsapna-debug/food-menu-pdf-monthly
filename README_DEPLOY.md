# Monthly Menu Web Portal - setup guide (all free)

You get a website usable from any system: pick month/year/holidays,
download the PDF, and edit the menu (week1..week5) in the browser.
The menu is stored in Supabase, the code in GitHub, the site on Vercel.

## What goes where

| Service (all free) | Holds | Cost |
|---|---|---|
| GitHub | code (`web/` folder) | free |
| Supabase | menu data (database) + `format.pdf` (storage) | free tier |
| Vercel | the website + PDF engine | free tier |

---

## OPTION A - full portal (GitHub + Supabase + Vercel)

### Step 1 - Supabase (database + artwork), ~10 min
1. Go to `supabase.com` -> sign up (GitHub login works) -> **New project**
   (name `school-menu`, free plan, any region close to you). Wait ~2 min.
2. Left menu **SQL Editor -> New query**, paste the whole `web/supabase.sql`,
   press **Run**. This creates the `app_data` table.
3. Left menu **Storage -> New bucket**, name it exactly `assets`,
   make it **Public**, Create.
4. Open the `assets` bucket -> **Upload** your `format.pdf`.
5. Left menu **Project Settings (gear) -> API**: copy and keep safely:
   - `Project URL` -> this is `SUPABASE_URL`
   - `service_role` secret key -> this is `SUPABASE_SERVICE_KEY`
     (never put this key in the browser/frontend - only in Vercel env vars).

### Step 2 - GitHub (code), ~5 min
1. In your repository, upload the **contents** of the `web/` folder:
   `index.html`, `vercel.json`, `requirements.txt`, `supabase.sql`,
   `api/menu.py`, `api/generate.py`
   (either to the repo root, or keep them under `web/` - see step 3 note).
2. Commit.

### Step 3 - Vercel (hosting), ~5 min
1. Go to `vercel.com` -> sign up with GitHub -> **Add New -> Project** ->
   **Import** your repository.
2. If the files are inside a `web/` subfolder: open **Build and Output
   Settings -> Root Directory -> Edit -> `web/`**.
3. **Environment Variables**: add both (paste from step 1.5):
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_KEY`
4. **Deploy**. You get a public link like
   `https://school-menu.vercel.app` - open it on any system.

### Step 4 - first use (one time)
1. Open your Vercel link -> section **2. Edit Menu** ->
   **Import menu.xlsx** (your existing file) -> **Save to Cloud**.
2. Section **1. Generate PDF**: pick month/year/holidays ->
   **Generate & Download PDF**.
3. From now on, editing happens in section 2 (per-week tables), or
   re-import an updated `menu.xlsx` / export one for records.

### If something fails
- "Cloud menu is empty" -> do Step 4.1 (Import + Save) first.
- Vercel deploy error -> check Root Directory and that both env vars exist
  (Project -> Settings -> Environment Variables), then Redeploy.
- PDF has no background/logos -> the `assets/format.pdf` upload (Step 1.3-1.4).

---

## OPTION B - simpler alternative (Streamlit only, no Supabase/Vercel)

Best when you just need "open link, upload xlsx, download PDF".

1. In your GitHub repo, upload `menu_app.py`, `format.pdf` and
   `web/streamlit_app.py` (rename it to `streamlit_app.py` at repo root
   if you like), plus a `requirements.txt` containing:
   ```
   streamlit
   openpyxl
   pypdf
   reportlab
   Pillow
   ```
2. Go to `share.streamlit.io` -> sign in with GitHub ->
   **New app** -> pick repo/branch/file -> **Deploy** (free).
3. Open the given link anywhere: upload `menu.xlsx` in the browser,
   enter month/year/holidays, download the PDF. Menu changes = edit the
   xlsx and re-upload. (No in-browser cell editing in this option.)

---

## Security notes
- The Supabase `service_role` key lives only in Vercel's server settings;
  visitors of the site can generate PDFs and edit the menu but never see it.
- Want the site private? Vercel Hobby: Project -> Settings -> Deployment
  Protection -> Vercel Authentication (visitors must log in), free.
