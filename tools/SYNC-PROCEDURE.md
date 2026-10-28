# Sync procedure (used by the scheduled daily sync)

Rebuilds all dashboards from the two Google Sheets and pushes them to `main`.

Sources (Google Drive file IDs)
- RM_WH_Mgt: `1wFwiKAfhgkTKgL3QvKj8yn2jWyxGw55infzXfJtZXsI`
- SBO Purchase Blanket Agreement: `1uavWZvYKFIivSr4Sz8NXDUdDZuQsSNH-PI6wH3LF4kQ`
- Inventory Price (an IMPORTRANGE copy of the Inventory (Ageing) tab of SBO Inventory Report; batch cost per item and warehouse): `17dwpOMvXaS3Rx1jmpX9Hf18X-74JoOLty5EapFxnbFY`

Steps
1. Download all three files with the Google Drive `download_file_content` tool using `exportMimeType` `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`. The result is too large to print and is saved to a file as JSON `{content: <base64>, title, ...}`.
2. In a scratch work folder, create `work/src/` and write the decoded bytes to `work/src/wh.xlsx` (RM_WH_Mgt), `work/src/pba.xlsx` (SBO Purchase Blanket Agreement) and `work/src/audit.xlsx` (Inventory Price). Check the `title` field to tell them apart. The Ageing copy's tab is called `Sheet1` (it may also have an `Ageing` tab; the build reads `Sheet1`).
3. From the repo root run (use `python3 -I`):
   - `python3 -I tools/build_data.py <work>/`   prints the sheets' own sync times and totals (the Inventory Price sync time is in data.json as `costSyncFull`)
   - `python3 -I tools/gen.py <work>/ ./`   writes index.html, packaging.html, micro.html, feedtags.html
   - `python3 -I tools/sync_page.py <work>/ ./ <today YYYY-MM-DD>`   writes sync.html and updates history.json
4. Sanity check: each page has the expected item counts and no errors (open with Playwright/Chromium if available).
5. If the INV, PBA and Inventory Price sync times are the same as the latest entry in `history.json` and today's snapshot adds no change, do not commit. Otherwise commit and push to `main`.

Notes
- The Sync tab keeps one snapshot per day in `history.json`. Running twice in a day replaces that day's snapshot.
- Do not edit the generated HTML by hand. Change `tools/template.html` or the Python scripts and rebuild.
- Expect a few seconds for the build; the downloads are the slow part.
