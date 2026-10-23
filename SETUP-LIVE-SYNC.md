# Live sync setup (Apps Script)

The **Sync now** button on each dashboard calls a small Apps Script that reads the two source sheets and returns fresh data. It is read-only and needs a secret key.

1. Go to https://script.google.com and create a **New project** while signed in as an account that can open both sheets (RM_WH_Mgt and SBO Purchase Blanket Agreement).
2. Replace the code with `apps-script/Code.gs` from this repo and save.
3. **Project Settings > Script properties > Add**: `ADMIN_KEY` = a long random string (16+ characters). Keep it secret.
4. **Deploy > New deployment > Web app**: Execute as **Me**, Who has access **Anyone**. Approve the permission prompt (the script reads your spreadsheets). Copy the `/exec` URL.
5. Open `config.js` in this repo and paste the URL into `window.SYNC_URL`, then commit.
6. Open a dashboard, click **Sync now**, and enter the key when asked. It is kept only for that browser tab. You can also open `index.html#key=YOUR_KEY` once; the key is removed from the address bar.

Notes
- The data returned includes stock, suppliers and prices. Anyone with the key can read it, so share the key like a password. Change it in Script properties to cut off old links.
- The script runs as you, so it only works while your account keeps access to both sheets.
- Sync now refreshes the page in your browser only. It does not change the saved snapshot or the history on the Sync tab.
- If a sheet's tab or column names change (INV, PBA, OPEN PBA + PO, MIN_* sheets, `AV W Fcasts consump`), the script will report an error instead of wrong numbers.
