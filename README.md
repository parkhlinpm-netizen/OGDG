# OGDG website (built from the Excel master)

The website is **generated** from `data/master.xlsx` by `scripts/make_ogdg_site.py`.
Nothing in `public/` is edited by hand. GitHub Actions rebuilds and publishes the site automatically.

## Update the website
1. Open the `data` folder, then **Add file > Upload files**.
2. Upload the new Excel master **with the file name `master.xlsx`** (same name = it replaces the old one).
3. Click **Commit changes**.
4. Open the **Actions** tab. A run called "Build and deploy the website" starts by itself (about 1-2 minutes).
   Green check = published. Red X = open the run to read the error.
5. The site is at `https://<account>.github.io/<repository-name>/` (the address does not change).

You can also start a build by hand: **Actions > Build and deploy the website > Run workflow**.

## Report PDF (optional)
Put the report PDF in the `report` folder. It is built into the site (the **Report** button).
Without a PDF the button is not shown.

## One-time setup
**Settings > Pages > Build and deployment > Source: GitHub Actions**.

## Rules that keep the links stable
- Keep the repository name and the account the same (they are part of the address).
- Never change an existing ID in the workbook (for example ADM-001). New items get new IDs.
- The sheets read are `Datasets`, `Use_Cases`, `Themes` and `Carinal 4`. Do not rename them.

## Run locally (optional)
```
pip install -r requirements.txt
python scripts/make_ogdg_site.py data/master.xlsx public/index.html [report/your.pdf]
```
Then open `public/index.html` in a browser.
