"""
NSE Nifty Indices - Streamlit Community Cloud connectivity spike test.

PURPOSE: this is intentionally minimal. Before building the full Streamlit
UI, we need to confirm niftyindices.com's Akamai bot-management doesn't
block requests coming from Streamlit Community Cloud's outbound (cloud
datacenter) IP ranges -- the desktop .exe this is replacing runs from a
residential/office IP, which is a different risk profile.

Deploy ONLY this file to Streamlit Community Cloud and check what it shows:
  - A populated list of sub-index types -> real data got through, proceed
    with the full port.
  - An error / raw HTML snippet -> requests are being blocked as a bot;
    stop and report back before investing more time in the full UI.

See STREAMLIT_MIGRATION_HANDOVER.md section 6 for full context.
"""

import streamlit as st

from nse_client import NSEClient

st.set_page_config(page_title="NSE API connectivity spike test", page_icon="[test]")

st.title("NSE API connectivity spike test")
st.caption(
    "Minimal probe to check whether niftyindices.com's Akamai bot-management "
    "blocks requests from this cloud IP. Not the real app."
)

if st.button("Run test: fetch Equity sub-index types"):
    with st.spinner("Calling niftyindices.com ..."):
        try:
            client = NSEClient()
            sub_types = client.get_sub_index_types("Equity")
        except Exception as e:
            st.error("Request failed or was blocked.")
            st.exception(e)
        else:
            if sub_types:
                st.success(f"Got {len(sub_types)} sub-index type(s) back — looks like real data.")
                st.write(sub_types)
            else:
                st.warning(
                    "Request succeeded but returned an empty list. "
                    "That's unexpected — worth a closer look before proceeding."
                )
