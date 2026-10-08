# Cashier Slip Dashboard

Pick a branch → see each cashier's total bet slips → click a cashier for their bet slips per game. Excel download at the end; upload page for new CSVs; data stored in Neon.

## Files
- `app.py` — the dashboard and upload pages
- `slipdb.py` — CSV cleaning, Neon tables, Excel export
- `requirements.txt`
- `.streamlit/secrets.toml.example` — where the Neon connection string goes

## Set up (Streamlit Community Cloud)
1. Push this folder to a GitHub repo (or a subfolder of your existing one).
2. In Streamlit Cloud, create an app pointing at `app.py`.
3. In **App settings → Secrets** paste:
   ```
   DATABASE_URL = "postgresql://USER:PASSWORD@ep-xxxx-pooler.REGION.aws.neon.tech/neondb?sslmode=require"
   ```
   (Neon console → Connect → copy the connection string.)
4. Open the app → **Upload data** → drop in all the monthly Slip Summary CSVs → **Save**.

The tables (`cs_slip_summary`, `cs_uploads`) are created automatically on first run. The `cs_` prefix keeps them apart from your existing dashboard's tables in the same Neon database.

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```
Without a `DATABASE_URL` it uses a local `local_dev.db` file for testing.

## Excel download
Contains what is on screen: the branch's cashiers for the chosen month plus the all-months total, and the clicked cashier's per-game breakdown.
