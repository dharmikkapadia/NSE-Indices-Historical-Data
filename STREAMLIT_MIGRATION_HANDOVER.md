# Handover: Port NSE Nifty Indices Downloader to Streamlit + deploy on Streamlit Community Cloud

**Origin:** this migration was scoped in a claude.ai chat (not visible to you). This doc is the
complete, self-contained handover of everything decided/discovered there. If something below is
ambiguous, ask the user rather than guessing — don't assume you can re-derive it from first
principles, especially the API section, which was reverse-engineered against the live site.

**Bring these 4 files into this session/repo before starting** (the user has them from the
claude.ai chat): `nse_downloader.py`, `requirements.txt`, `build_exe.bat`, `README.md`.

---

## 1. What this app is today

A Windows desktop app (CustomTkinter/Tkinter, compiled to `.exe` via PyInstaller) for
**Arthashastra Finsec Providers Pvt Ltd (AFP)**. It downloads historical OHLC data for selected
NSE Nifty indices from niftyindices.com and writes one `.xlsx` workbook per run — one sheet per
selected index, plus an optional combined "All Indices" sheet. Fixed Sepia (warm parchment) color
theme, no theme switcher.

## 2. Goal

Rebuild the UI layer as a Streamlit app and deploy it on Streamlit Community Cloud, so AFP
research desk users can run this from a browser instead of a Windows `.exe`.

## 3. Why this isn't a "just deploy it" task

Streamlit Community Cloud runs a Python script that imports `streamlit` and renders through
Streamlit's own widget model, in a headless container with no display server. The current app's
UI is CustomTkinter/Tkinter, which requires an actual display to draw windows — it cannot run
inside Streamlit's container at all. **This is a UI rewrite, not a deployment step.**

## 4. What's fully reusable, unchanged (zero Tkinter dependency)

These live in `nse_downloader.py` and can be copied verbatim into the new Streamlit app (put them
in a separate module, e.g. `nse_client.py`, imported by the Streamlit entrypoint):

- `NSEClient` class — all three API methods (`get_sub_index_types`, `get_indices`,
  `get_historical_data` / `_fetch_chunk`), plus `_parse_json_response`
- `build_catalog(client, log_fn)` — builds the full index catalog across Equity / Fixed Income /
  Multi Asset
- `safe_sheet_name(name, used_names)`, `_to_num(v)`, `_style_header`, `_write_index_sheet`,
  `_write_combined_sheet` — the openpyxl workbook-writing helpers
- `parse_date_str(s)`
- `BASE_URL`, `HEADERS`, `INDEX_TYPES`, `MAX_CHUNK_DAYS` constants

None of this touches `customtkinter`, `tkinter`, threading, or the filesystem in a way that's
Windows-specific. It's straight `requests` + `openpyxl`.

## 5. CRITICAL — current live API facts (hard-won; do not re-derive from scratch)

niftyindices.com migrated its backend mid-2026 from an old ASP.NET WebForms pattern
(`/Backpage.aspx/{method}`) to a new one (`/BackPage/{method}`). This was found by intercepting
real browser traffic against the live site, not by guessing. If you need to touch the API layer,
these are the current facts as of this handover:

| Endpoint | Method | Payload (`json=`) | Response |
|---|---|---|---|
| `{BASE_URL}/BackPage/gethistoricaltypeSubindexdata` | POST | `{"cinfo": {"indextype": <str>, "indexgroup": "Historical Index Data"}}` | Direct JSON array of objects with an `"indextype"` key. **No `.d` wrapper** (that was the old ASMX pattern). |
| `{BASE_URL}/BackPage/gethistoricaltypeindexdata` | POST | `{"cinfo": {"indextype": <sub_type>, "indexgroup": "Historical Index Data"}}` | Same shape as above. |
| `{BASE_URL}/BackPage/getHistoricaldatatabletoString` | POST | `{"cinfo": "<single-quoted pseudo-JSON string>"}` — **note this one's `cinfo` is a string, not a nested object**: `"{'name':'<index>','startDate':'DD-Mon-YYYY','endDate':'DD-Mon-YYYY','indexName':'<index>'}"` | Direct JSON array of records: `RequestNumber, Index Name, INDEX_NAME, HistoricalDate ("DD Mon YYYY" string), OPEN, HIGH, LOW, CLOSE` (all values as strings — `_to_num()` handles conversion). |

Other facts:
- The site enforces roughly a ~1 year window per historical-data request; `NSEClient` chunks
  longer ranges into `MAX_CHUNK_DAYS = 360`-day windows and merges/dedupes by date.
- A session GET to `{BASE_URL}/reports/historical-data` is done once at `NSEClient.__init__` to
  seed cookies before any POST — keep this.
- Required headers: `Content-Type: application/json; charset=UTF-8`, `Accept`,
  `X-Requested-With: XMLHttpRequest`, a real desktop Chrome `User-Agent`, `Referer` and `Origin`
  both set to `https://www.niftyindices.com` (see `HEADERS` dict in the source).
- If the site ever changes again, the symptom will be `_parse_json_response` raising a
  `RuntimeError` with a snippet of the non-JSON body it got back (this was added specifically so
  a future break is self-diagnosing instead of a cryptic `JSONDecodeError`).

## 6. KNOWN RISK — test this *first*, before building the full UI

niftyindices.com has **Akamai bot-management** active — its sensor beacon
(`/akam/13/{id}`) was observed live on the page. Cloud datacenter IP ranges (which is what
Streamlit Community Cloud's outbound traffic comes from) are exactly the kind of traffic Akamai
bot-management tends to flag more aggressively than a residential/office IP, which is what the
current desktop `.exe` runs from. **This was never tested from a cloud IP** — it's an open risk,
not a confirmed blocker.

**Do this before writing any Streamlit UI code:**
1. Deploy a *minimal* one-file Streamlit app to Community Cloud that does nothing but
   `NSEClient().get_sub_index_types("Equity")` and displays the result.
2. If it returns real data → proceed with the full port below.
3. If it gets blocked / returns HTML / raises the `RuntimeError` from `_parse_json_response` →
   stop and flag this to the user before investing more time. Fallback options to discuss with
   them at that point: a scheduled scrape (e.g. GitHub Actions or a small VM with a stable IP)
   writing to a lightweight datastore that the Streamlit app just reads from, rather than the web
   app hitting niftyindices.com directly per-request.

## 7. UI migration map

| Current (CustomTkinter) | Streamlit equivalent |
|---|---|
| Custom collapsible category/index checkbox tree (`IndexListWidget`) | `st.multiselect` per category, or a simpler flat searchable `st.multiselect` across all indices — doesn't need to replicate the collapsible tree exactly, use judgment on what's usable in Streamlit's widget set |
| Filter text box | Native to `st.multiselect` (built-in type-to-filter) |
| Start/End date entries + quick-set buttons (Prev Month / This Month / YTD / Last 1Y) | `st.date_input` (range or two inputs) + `st.button` per quick-set that mutates `st.session_state` before the date inputs render |
| Output folder picker (`filedialog.askdirectory`) | **Removed entirely** — see §8 |
| "Also include All Indices combined sheet" checkbox | `st.checkbox` |
| Save Preset / Load Preset / Clear buttons | `st.session_state` + optionally `st.cache_data`-backed persistence if presets should survive across sessions (see §9 on persistence caveats) |
| Progress bar + status dot + Activity log textbox | `st.progress` + `st.status` / `st.empty()` updated in a loop |
| DOWNLOAD button | `st.download_button`, fed by the in-memory workbook (§8) |
| Sepia palette (`P` dict in the source) | `.streamlit/config.toml` `[theme]` section using the same hex values, or custom CSS via `st.markdown(unsafe_allow_html=True)` if finer control is needed than the theme config allows |

## 8. Output model change — this is the biggest structural change

The desktop app writes to a user-chosen local folder via `filedialog.askdirectory`. That concept
doesn't exist in a browser app — there's no "the server's folder" the user can browse to
meaningfully. Build the workbook **in memory** and serve it for download:

```python
import io
buf = io.BytesIO()
wb.save(buf)
buf.seek(0)
st.download_button(
    "Download workbook",
    data=buf,
    file_name=f"NSE_Indices_{start.isoformat()}_to_{end.isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
```

Convenient fact: the desktop app was already redesigned (in the same claude.ai session) to
produce **one fresh, standalone workbook per run with no merging across runs** — the user
explicitly chose that model over a persistent/growing workbook. That maps cleanly onto a
stateless download-button pattern; no incremental-merge logic to port.

## 9. Config / state / caching changes

- `preset_indices`, `output_folder`, `combined_sheet` were persisted to a local
  `nse_downloader_config.json` next to the `.exe`. On Community Cloud, the container filesystem
  is **ephemeral and shared across every visitor** — there is no per-user local file. Use
  `st.session_state` for in-session persistence. If presets need to survive across browser
  sessions/devices, that requires an actual backend (e.g. a small SQLite/Supabase/gspread store)
  — flag this as a scope question for the user rather than assuming it's needed.
- The index catalog (`indices_catalog.json`, cached locally so subsequent launches are instant)
  should become `@st.cache_data(ttl=...)` (pick a sensible TTL, e.g. 24h) wrapping
  `build_catalog()`. Keep `NSEClient` itself in `@st.cache_resource` (one shared `requests.Session`
  per app instance, not rebuilt every rerun).

## 10. Threading/progress model change

The desktop app uses a background `threading.Thread` + `queue.Queue` to keep the Tkinter event
loop responsive during downloads (`_download_thread`, `_poll_queue`, `_log_async`, etc.). Streamlit
reruns the whole script top-to-bottom on each interaction and has its own idioms for this —
`st.spinner`, `st.progress(i / n_total)` updated in a plain loop, `st.status()` as a log container.
**None of the threading/queue machinery needs to be ported** — it can be dropped entirely, which
significantly shrinks the code versus a literal translation.

## 11. Styling reference — Sepia palette (carry over if visual parity matters)

```
bg            #F3E9D2
surface       #FBF4E4
surface_2     #EFDFBF
surface_hover #E5D1A3
border        #D6BE8E
text          #4A3728
text_muted    #8A7256
primary       #8B5A2B   (sepia brown — primary accent / buttons)
primary_hover #6E4720
secondary     #B08442   (amber / gold)
success       #6B8E4E
warning       #C08A2E
error         #A24936
```

## 12. Suggested build order

1. **Akamai spike test first** (§6) — do not skip this, it determines whether the rest is worth doing as designed.
2. Repo structure: `streamlit_app.py` (entrypoint), `nse_client.py` (ported logic from §4),
   `requirements.txt` (`streamlit`, `requests`, `openpyxl` — no `customtkinter`/`darkdetect`,
   no `pyinstaller`), `.streamlit/config.toml` (Sepia theme).
3. Minimal working UI: catalog load → multiselect → date range → download button producing a
   correct in-memory `.xlsx`. Get this correct before layering on polish.
4. Add caching (`st.cache_data` for catalog, `st.cache_resource` for `NSEClient`).
5. Add Sepia styling / layout polish.
6. Test locally: `streamlit run streamlit_app.py`.
7. Push to GitHub, deploy via share.streamlit.io (Community Cloud reads `requirements.txt`
   automatically from the repo root).
8. Configure access control — see open question below.

## 13. Open questions to confirm with the user (don't assume)

- **Access control**: should the deployed app be public, or gated behind Community Cloud's
  email allow-list, given this is AFP-internal tooling? This was raised but not yet answered as
  of this handover.
- Keep the desktop `.exe` build available alongside the web version, or fully replace it?
- Does preset persistence need to survive across browser sessions/devices (→ needs a real
  backend), or is per-session (`st.session_state`) sufficient?
