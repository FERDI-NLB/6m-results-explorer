# 6 M Results Explorer — automatic updates

This repository includes the current standalone webpage, original Excel dataset, and an hourly GitHub Actions workflow.

## Publish

1. Create a new **public GitHub repository** (e.g. `6m-results-explorer`).
2. Upload **the contents** of this folder to the repository root, including the hidden `.github/workflows/update.yml` file. The simplest way to preserve hidden files is to extract the ZIP and push with Git, or use GitHub Desktop.
3. In repository **Settings → Pages → Build and deployment**, choose **GitHub Actions** as the source.
4. In repository **Settings → Actions → General → Workflow permissions**, allow **Read and write permissions** if needed.
5. Open **Actions → Publish website → Run workflow** to deploy the initial site. Then open **Actions → Verify official results and publish → Run workflow** to test the updater. Your public address will be `https://YOUR-USERNAME.github.io/YOUR-REPOSITORY/`.
6. GitHub runs the checker at minute 17 of each hour (actual start may be delayed). Successful verified changes commit the updated Excel workbook and HTML, and triggers the separate publishing workflow to redeploy the public site.

## Important — official source verification required

The checker intentionally **fails closed** until its official NLB HTML parser has been tested against the *actual live page structure*. The provided URL currently cannot be inspected reliably from this environment. Do not remove this safeguard or assume automatic result ingestion works without checking the Actions logs. It will **never fabricate a draw**, scrape unofficial sites as if official, or overwrite an existing result that conflicts.

The workbook lists draw 6348 as **D786987**, whereas an NLB-branded listing currently displays **D789687**. Verify the discrepancy against the official record before relying on this dataset. Baseline 6348 is left unchanged by the updater.

## Workbook and website behavior

- `data.xlsx` is the canonical data file **in the GitHub repository** once published. A local copy on your computer will not automatically synchronize. Download the latest workbook from GitHub or clone/pull the repository to keep your computer in sync.
- `docs/index.html` contains embedded data. The updater regenerates it after successful result ingestion. Predictions recalculate when visitors load the updated page.
- Existing blank future rows are filled by draw number; repeated runs are idempotent.
- If official NLB HTML changes, parsing fails and **nothing is published**. Review the Actions log and adapt the parser to the site's actual markup.
- This does not bypass CAPTCHAs, access restrictions, or website policies. Respect official site terms and rate limits.

## Technical validation

Run `python scripts/update.py` locally after installing `requirements.txt`. A successful run with no new result makes no changes. An official result must include draw number, ISO date, letter and exactly six digits in the same result block. Conflicts, missing intermediate draws, future dates, or unreachable source cause the workflow to fail.
