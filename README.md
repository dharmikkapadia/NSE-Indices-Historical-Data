# NSE Nifty Indices  -  Historical Data Downloader

A modern Windows desktop app for **Arthashastra Finsec Providers Pvt Ltd (AFP)**
that downloads historical OHLC data for selected NSE Nifty indices from
[niftyindices.com](https://www.niftyindices.com/reports/historical-data) into
a single Excel workbook -- one sheet per index.

Built with `customtkinter` for a clean, modern look with **System / Light / Dark**
theme support. Compiles to a single `.exe` that runs on **any Windows PC with
no Python installation required**.

---

## Files in this folder

| File                  | Purpose                                                   |
|-----------------------|-----------------------------------------------------------|
| `nse_downloader.py`   | The application source (Tkinter GUI + API client).        |
| `build_exe.bat`       | Double-click to build `NSEDataDownloader.exe`.            |
| `requirements.txt`    | Python dependencies (used by `build_exe.bat`).            |
| `README.md`           | This file.                                                |

After building, two extra files are created at runtime next to the .exe:

| File                            | Purpose                                                  |
|---------------------------------|----------------------------------------------------------|
| `indices_catalog.json`          | Cached list of all available indices.                    |
| `nse_downloader_config.json`    | Saved preset of selected indices + last output folder.   |

---

## Building the .exe (one-time setup)

You only need to do this on **one** PC that has Python installed.
The resulting `.exe` is portable.

1. Install Python 3.9 or newer from <https://www.python.org/downloads/>.
   **Tick "Add python.exe to PATH"** during installation.
2. Double-click `build_exe.bat` (or run it from a Command Prompt).
3. Wait 1-3 minutes. When done, the executable will be at:
   ```
   dist\NSEDataDownloader.exe
   ```
4. Copy that single `.exe` to any Windows PC and double-click to run.

> The `.bat` automatically installs `pyinstaller`, `requests`, `certifi`,
> `customtkinter`, `darkdetect`, and `openpyxl`.

---

## Using the app

### First launch
On first launch the app fetches the full index catalog from niftyindices.com
(takes about 5-10 seconds) and caches it locally as `indices_catalog.json`.
Subsequent launches are instant.

### Workflow

1. **Select indices** in the left panel. Click a category row to toggle
   all indices under it; click a single index to toggle just that one.
   Use the **Filter** box to find an index by name.
2. **(Optional) Save selection as preset** -- the next time you open the
   app, your selection is restored automatically.
3. **Pick a date range** (YYYY-MM-DD format).
   Quick-set buttons: *Prev Month*, *This Month*, *YTD*, *Last 1Y*.
4. **Choose an output folder.**
5. **Click DOWNLOAD.**

### Theme

The **Theme** dropdown in the top-right toolbar lets you switch between:

- **System** *(default)* -- automatically matches the Windows light/dark setting.
- **Light** -- always light.
- **Dark** -- always dark.

The choice is saved and restored on next launch.

### Output structure

Each download produces **one Excel workbook** in your output folder, covering
the entire date range you picked (no more splitting by month):

```
<Output Folder>/
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

Each run is a **fresh, standalone file** named after the date range you
entered -- it does not merge with or read from any previous download. If
you re-run the same date range, that file is overwritten.

### Behavior options

- **Also include an "All Indices" combined sheet** -- adds the long-format
  sheet described above on top of the per-index sheets. Turn it off if you
  only want the individual index sheets.

---

## Recurring downloads (e.g. a monthly pull)

Each run makes its own dated file, so nothing gets merged automatically:

1. Open `NSEDataDownloader.exe`.
2. Click **Load Preset** (your saved indices auto-tick).
3. Pick the date range you want (e.g. **Prev Month**, or a custom range).
4. Click **DOWNLOAD**.

That writes a workbook named for that date range, e.g.
`NSE_Indices_2026-06-01_to_2026-06-30.xlsx`, alongside any earlier
downloads in the same output folder -- each file is independent.

---

## Troubleshooting

| Issue                              | Fix                                                                                         |
|------------------------------------|---------------------------------------------------------------------------------------------|
| `python` not recognized            | Re-install Python with **Add to PATH** ticked, or reboot.                                  |
| Build fails on `pyinstaller`       | Run as Administrator, or `python -m pip install --upgrade pyinstaller` manually.            |
| Build fails on `customtkinter` install | If you're on a brand-new Python release (e.g. 3.14) and pip can't find a compatible `customtkinter`, install Python 3.12 alongside and run `build_exe.bat` from that. |
| Catalog fetch fails                | Check internet / proxy / firewall. Click **Refresh Catalog** to retry.                      |
| Antivirus flags the .exe           | Common false positive for PyInstaller bundles. Add an exception or rebuild on that PC.      |
| Some indices return 0 records      | The index either did not exist in that date range, or the site returned an empty payload. Its sheet is still created, with headers only. Try a wider range. |
| "Could not save ... close it in Excel first" | The output workbook is already open in Excel (or another program). Close it and click **DOWNLOAD** again. |
| Need data older than 1 year        | The app auto-chunks long ranges into yearly segments under the hood -- just enter the dates. |
| Theme stays Light when set to System | This means `darkdetect` couldn't read the Windows setting. Use the dropdown to force Light or Dark. |

---

## Technical notes

- **Tech stack**: Python 3, [`customtkinter`](https://github.com/TomSchimansky/CustomTkinter)
  (modern Tk widgets), `darkdetect` (system-theme detection), `requests`,
  `openpyxl` (xlsx workbook output), PyInstaller.
- **No browser automation, no Selenium, no Chrome dependency.** Pure HTTP
  against the public JSON endpoints used by the niftyindices.com page itself.
- **Threaded downloads** keep the UI responsive; a queue marshals log
  messages back to the main thread.
- **Date-range chunking**: the site's API caps each request at ~1 year, so
  longer ranges are split into 360-day chunks and re-merged.
- **Theming**: the **Theme** dropdown in the toolbar toggles between
  *System* (auto-follows the Windows colour-mode setting), *Light*, and
  *Dark*. Your choice is saved in `nse_downloader_config.json` and
  restored on next launch.

---

*Built for AFP -- internal research tooling.*
