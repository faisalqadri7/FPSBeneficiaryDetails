# FPS Beneficiary Details Web App

A static web app that automatically loads and displays data from `FPSBeneficiaryDetailsbalapora.xlsx` in the browser using SheetJS.

## Project files

- `index.html` - page structure and script/style wiring
- `style.css` - modern responsive styling
- `script.js` - Excel loading, parsing, table rendering, and live search
- `FPSBeneficiaryDetailsbalapora.xlsx` - source data file loaded on page open

## Run locally

Because browsers block file fetches from `file://` for many setups, run a local static server:

```bash
python3 -m http.server 8080
```

Then open:

```text
http://localhost:8080
```

## Push this project to GitHub

1. Create a new GitHub repository (for example: `fps-beneficiary-details`).
2. In this project folder, run:

```bash
git init
git add .
git commit -m "Create static FPS beneficiary viewer"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

> If this repository is already initialized, skip `git init` and just commit/push to the existing remote.

## Deploy on Cloudflare Pages

1. Log in to Cloudflare Dashboard.
2. Go to **Workers & Pages** → **Create application** → **Pages** → **Connect to Git**.
3. Select your GitHub repository.
4. Build settings for this static app:
   - **Framework preset:** `None`
   - **Build command:** *(leave empty)*
   - **Build output directory:** `/` (root)
5. Click **Save and Deploy**.
6. Cloudflare Pages will publish your site and provide a `*.pages.dev` URL.

## Updating the Excel data

1. Replace `FPSBeneficiaryDetailsbalapora.xlsx` in the repository root with the updated file.
2. Commit and push:

```bash
git add FPSBeneficiaryDetailsbalapora.xlsx
git commit -m "Update beneficiary details data"
git push
```

Cloudflare Pages auto-deploys after each push, and the app will load the updated Excel file automatically on next page load.
