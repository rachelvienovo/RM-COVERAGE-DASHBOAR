# RM Coverage Dashboards

Static dashboards showing days of cover for raw materials.

- `index.html`: Macro raw materials (tonnes)
- `packaging.html`: Packaging (pieces)
- `micro.html`: Micro ingredients (kg)
- `feedtags.html`: Feedtags (pieces)
- `sync.html`: data freshness, snapshot history and changes since the last sync (`history.json` holds the snapshots)

Data is a snapshot embedded in each page (as of 2026-10-20). Stock and consumption come from RM_WH_Mgt; pending comes from the OPEN PBA + PO tab of SBO Purchase Blanket Agreement.
