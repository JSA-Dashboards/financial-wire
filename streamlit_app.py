"""Morning Financial Wire - hosted view.

Shows today's wire (snapshot, reading list, macro read, history) and the
headline-candidate picker. Nothing here writes into the wire: ticking items only
builds a copy-ready bullet list for you to rewrite.

Email-sourced content arrives from Snowflake (push_wire.py on your PC). The
public news feeds are fetched live. With no Snowflake config it reads the local
files, so `streamlit run streamlit_app.py` works on your PC too.
"""
import streamlit as st

import build_fin_wire
import headlines
import wire_store

st.set_page_config(page_title="Morning Financial Wire", page_icon=":material/newspaper:", layout="centered")


@st.cache_data(ttl=300, show_spinner="Loading the wire...")
def load_payloads():
    return wire_store.read_latest()


@st.cache_data(ttl=900, show_spinner="Fetching news feeds...")
def load_feeds():
    return headlines.collect(include_mailbox=False)


st.title("Morning financial wire")

payloads, origin = load_payloads()
wire = (payloads.get("wire") or {}).get("data")

meta = payloads.get("wire") or {}
stamp = meta.get("pushed_at")
st.caption(f"Source: {origin}" + (f" · email items pushed {stamp:%a %b %d %H:%M}" if stamp else ""))

if not wire:
    st.warning("No wire data yet. Run push_wire.py on the PC that has Outlook.")
else:
    with st.container(border=True):
        st.iframe(build_fin_wire.build(wire), height=1400)
    for g in wire.get("gaps") or []:
        st.caption(f":material/warning: {g}")

st.header("Headline candidates")
st.caption("Tick what matters, copy the list, paste into the wire's Headlines box and rewrite it in your own words.")


@st.fragment
def picker():
    feed_items, feed_errors = load_feeds()
    items = list((payloads.get("pj") or {}).get("data") or []) + \
        list((payloads.get("bbg") or {}).get("data") or []) + feed_items
    if not items:
        st.info("No candidates yet.")
        return

    topics = list(dict.fromkeys(i["topic"] for i in items))
    picked = []
    for t in topics:
        group = [i for i in items if i["topic"] == t]
        with st.expander(f"{t} ({len(group)})", expanded=t in ("PJ Quaid", "Bloomberg Markets Daily")):
            for n, it in enumerate(group):
                label = it["title"] if len(it["title"]) < 400 else it["title"][:397] + "..."
                tail = f"  \n:gray[{it['source']} · {it['when']}]"
                link = f" [:material/open_in_new:]({it['link']})" if it.get("link") else ""
                if st.checkbox(label + link + tail, key=f"{t}-{n}"):
                    picked.append(it)

    st.subheader(f"Ticked ({len(picked)})")
    if picked:
        st.code("\n".join(f"• {p['title']} ({p['source']})" for p in picked), language=None, wrap_lines=True)
    for e in feed_errors:
        st.caption(f":material/warning: feed failed: {e}")


picker()
