# NSE Nifty Indices  -  Historical Data Downloader

A Streamlit web app for **Arthashastra Finsec Providers Pvt Ltd (AFP)** that
downloads historical OHLC data for selected NSE Nifty indices from
[niftyindices.com](https://www.niftyindices.com/reports/historical-data) into
a single Excel workbook -- one sheet per index -- and serves it straight to
your browser as a download. No install required; runs on
[Streamlit Community Cloud](https://streamlit.io/cloud).

Sepia (warm parchment) color theme.

---

## Files in this folder

| File                       | Purpose                                                     |
|-----------------------------|--------------------------------------------------------------|
| `streamlit_app.py`          | The app entrypoint -- UI, download workflow, workbook build. |
| `nse_client.py`             | API client (`NSEClient`) + openpyxl workbook helpers.        |
| `requirements.txt`          | Python dependencies for Streamlit Community Cloud.           |
| `.streamlit/config.toml`    | Sepia theme configuration.                                   |
| `README.md`                 | This file.                                                    |

There is no server-side output folder and no local config file: presets and
selections live in your browser session (`st.session_state`) and reset when
the session ends. The index catalog is cached server-side for 24 hours
(`st.cache_data`) so it doesn't need to be re-fetched on every page load.

---

## Running locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Opens at `http://localhost:8501`.

## Deploying to Streamlit Community Cloud

1. Push this repo to GitHub.
2. On [share.streamlit.io](https://share.streamlit.io), create a new app
   pointing at this repo, branch, and `streamlit_app.py` as the entrypoint.
   `requirements.txt` at the repo root is picked up automatically.
3. Since this is AFP-internal tooling, restrict access via Community Cloud's
   viewer email allow-list (App settings -> Sharing) rather than leaving the
   app public.

---

## Using the app

### First load
On first load the app fetches the full index catalog from niftyindices.com
(a few seconds) and caches it for 24 hours. Use **Refresh Catalog** to force
a re-fetch sooner.

### Workflow

1. **Select indices** -- expand a category and use its search-as-you-type
   multiselect. Selections across categories combine into one list.
2. **(Optional) Save Preset** -- stores your current selection for the rest
   of this browser session. **Load Preset** restores it; **Clear** empties
   the current selection. Presets do not persist across sessions/devices.
3. **Pick a date range** with the date pickers, or a quick-set button:
   *Prev Month*, *This Month*, *YTD*, *Last 1Y*.
4. **Click DOWNLOAD.** Progress and per-index status show live; when done, a
   **Download workbook** button appears with the finished `.xlsx` in memory.

### Output structure

Each run produces **one Excel workbook**, covering the entire date range you
picked, named after that range:

```
NSE_Indices_2026-02-01_to_2026-02-28.xlsx
```

Inside the workbook:

- **One sheet per selected index**, named after the index (truncated to
  Excel's 31-character sheet-name limit where needed), with columns
  `Date, Open, High, Low, Close` sorted chronologically. Dates and OHLC
  values are stored as real Excel date/number types, not text.
- **An optional `All Indices` sheet** (toggle in Options) with every
  selected index in one long table -- columns `Index, Date, Open, High,
  Low, Close`, sorted by date then index name. Handy for pivoting.

Each run is a **fresh, standalone file** -- nothing is merged with or read
from a previous download.

### Behavior options

- **Also include an "All Indices" combined sheet** -- adds the long-format
  sheet described above on top of the per-index sheets. Turn it off if you
  only want the individual index sheets.

---

## Troubleshooting

| Issue                              | Fix                                                                                         |
|------------------------------------|-----------------------------------------------------------------------------------------------|
| Catalog fails to load              | The site may be temporarily unavailable. Click **Refresh Catalog** to retry.                |
| Some indices return 0 records      | The index either did not exist in that date range, or the site returned an empty payload. Its sheet is still created, with headers only. Try a wider range. |
| Need data older than 1 year        | Long ranges are auto-chunked into 360-day segments under the hood -- just enter the dates.  |
| Non-JSON / bot-check error         | niftyindices.com's API may have changed shape again, or requests from this app's IP are being blocked as a bot. The error message includes a snippet of what came back for diagnosis. |
| Presets disappeared                | Presets live in browser session state only -- they reset when the browser tab/session ends. |

---

## Technical notes

- **Tech stack**: Python 3, `streamlit`, `requests`, `openpyxl` (xlsx
  workbook output). No browser automation, no Selenium, no Chrome
  dependency -- pure HTTP against the public JSON endpoints used by the
  niftyindices.com page itself.
- **In-memory workbook**: built with `openpyxl` into a `BytesIO` buffer and
  served via `st.download_button` -- nothing is written to the server's
  filesystem, which is ephemeral and shared across every visitor on
  Community Cloud.
- **Date-range chunking**: the site's API caps each request at ~1 year, so
  longer ranges are split into 360-day chunks and re-merged.
- **Caching**: `NSEClient` (one shared `requests.Session`) is a
  `st.cache_resource`; the index catalog is `st.cache_data` with a 24h TTL.

---

*Built for AFP -- internal research tooling.*
