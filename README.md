# RM Coverage Dashboards

Static dashboards showing days of cover for raw materials.

- `index.html`: Macro raw materials (tonnes)
- `packaging.html`: Packaging (pieces)
- `micro.html`: Micro ingredients (kg)
- `feedtags.html`: Feedtags (pieces)
- `sync.html`: data freshness, snapshot history and changes since the last sync (`history.json` holds the snapshots)
- `tools/`: the build scripts and `SYNC-PROCEDURE.md` (how the dashboards are rebuilt from the sheets)

Data is a snapshot embedded in each page (as of 2026-10-20). Stock and consumption come from RM_WH_Mgt; pending comes from the OPEN PBA + PO tab of SBO Purchase Blanket Agreement.

## Live sync

Each dashboard has a **Sync now** button that refreshes the page from the Google Sheets through an Apps Script. See `SETUP-LIVE-SYNC.md`; the script is `apps-script/Code.gs` and the URL goes in `config.js`.
