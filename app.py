"""
Antidote dashboard.

WHAT THIS IS
------------
A live web page for the business. It reads the tracker spreadsheet, recomputes
every number from scratch, and shows the deck alongside it. One URL for a
recruiter: numbers on the left, slides on the right, deck downloadable.

HOW TO RUN IT ON YOUR OWN COMPUTER
-----------------------------------
1.  pip install streamlit openpyxl matplotlib
2.  Put this file, the tracker, antidote_analysis.py, and the slides/ folder
    in the SAME folder.
3.  In that folder run:   streamlit run app.py
4.  It opens in your browser.

HOW TO PUT IT ONLINE (one URL for the resume)
----------------------------------------------
1.  Make a free GitHub account.
2.  Upload this whole folder to a new repository.
3.  Go to share.streamlit.io, sign in with GitHub, pick the repo, pick app.py.
4.  It gives you a public link. Add rows to the tracker, push, page updates.

WHAT TO EDIT LATER
------------------
Nothing here holds numbers. Everything comes from the tracker through
antidote_analysis.py. To change a category or a forward assumption, edit the
CONFIG block at the top of antidote_analysis.py, not this file.
"""

import os
import streamlit as st

import antidote_analysis as aa

st.set_page_config(page_title="Antidote", page_icon="🟡", layout="wide")

YEL = "#FFC72C"
DARK = "#1C1C1A"

st.markdown(f"""
<style>
  html, body, [class*="css"] {{ font-family: "Times New Roman", Times, serif; }}
  .stApp {{ background: #FFFFFF; }}
  h1, h2, h3 {{ font-family: "Times New Roman", Times, serif; color: {DARK}; }}
  .rule {{ border-top: 6px solid {YEL}; width: 72px; margin: 0 0 14px 0; }}
  .sub {{ color: #6E6E68; font-style: italic; font-size: 1.05rem; }}
  [data-testid="stMetricValue"] {{ font-family: "Times New Roman", serif; color: {DARK}; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load(_tracker_mtime):
    """Read the tracker, recompute every number, and redraw every chart.
    The _tracker_mtime argument makes the cache refresh whenever the
    spreadsheet changes, so new sales show up without touching this file."""
    rows = aa.load_rows(aa.INPUT_FILE)
    per_line, _ = aa.summarize(rows)          # prints to server logs only
    bands = aa.ticket_band_analysis(rows)
    aa.make_charts(rows, per_line, bands)     # redraws the c_*.png charts
    sold = [x for x in rows if x["status"] == "Sold"]
    inv = [x for x in rows if x["status"] == "Inventory"]
    fluke = [x for x in sold if x["category"] == "Fluke test equipment"]
    fluke_paid = [x for x in fluke if x["acq"] == "Paid"]
    d = {
        "n_sold": len(sold),
        "proceeds": sum(x["net"] for x in sold),
        "profit": sum(x["profit"] for x in sold),
        "cogs": sum(x["cogs"] for x in sold),
        "inv_cash": sum(x["cogs"] for x in inv),
        "n_inv": len(inv),
        "paid_profit": sum(x["profit"] for x in sold if x["acq"] == "Paid"),
        "free_profit": sum(x["profit"] for x in sold if x["acq"] == "Free"),
        "fluke_profit": sum(x["profit"] for x in fluke),
        "fluke_cogs": sum(x["cogs"] for x in fluke_paid),
        "fluke_n": len(fluke),
        "fluke_days": (sum(x["turnaround_days"] for x in fluke if x["turnaround_days"])
                       / max(1, len([x for x in fluke if x["turnaround_days"]]))),
        "all_days": (sum(x["turnaround_days"] for x in sold if x["turnaround_days"])
                     / max(1, len([x for x in sold if x["turnaround_days"]]))),
    }
    d["margin"] = d["profit"] / d["proceeds"] if d["proceeds"] else 0
    d["fluke_roi"] = d["fluke_profit"] / d["fluke_cogs"] if d["fluke_cogs"] else 0
    d["fluke_share"] = d["fluke_profit"] / d["profit"] if d["profit"] else 0
    d["repeat_share"] = d["paid_profit"] / d["profit"] if d["profit"] else 0
    return d, rows


def slide(n):
    """Show one deck slide if the image is there."""
    p = f"slides/slide-{n:02d}.png"
    if os.path.exists(p):
        st.image(p, use_container_width=True)
    else:
        st.info(f"Slide {n} image not found. Export the deck to slides/ first.")


def chart(name, caption=None):
    if os.path.exists(name):
        st.image(name, use_container_width=True)
        if caption:
            st.caption(caption)


try:
    D, ROWS = load(os.path.getmtime(aa.INPUT_FILE))
except Exception as e:
    st.error(f"Could not read the tracker. Check INPUT_FILE in antidote_analysis.py.\n\n{e}")
    st.stop()

# ---------------- header ----------------
st.markdown('<div class="rule"></div>', unsafe_allow_html=True)
st.markdown("# Antidote")
st.markdown('<p class="sub">Small, certain, fast.</p>', unsafe_allow_html=True)
st.write(
    "I buy underpriced test equipment on Facebook Marketplace and resell it on eBay. "
    "This page reads my own transaction records and recomputes every figure below. "
    "Nothing is typed in by hand."
)

c = st.columns(4)
c[0].metric("Sales", f"{D['n_sold']}")
c[1].metric("Money in", f"${D['proceeds']:,.0f}")
c[2].metric("Profit", f"${D['profit']:,.0f}")
c[3].metric("Return on Fluke", f"{D['fluke_roi']*100:.1f}%")
st.caption(
    f"${D['inv_cash']:,.0f} sits in {D['n_inv']} unsold units. That is inventory at cost, not a loss."
)

if os.path.exists("Antidote_Review_and_Fluke_Case_Study.pptx"):
    with open("Antidote_Review_and_Fluke_Case_Study.pptx", "rb") as f:
        st.download_button("Download the deck (PowerPoint)", f,
                           file_name="Antidote_Review_and_Fluke_Case_Study.pptx")

tab1, tab2, tab3 = st.tabs(["Profitability review", "Fluke case study", "The deck"])

# ---------------- tab 1 ----------------
with tab1:
    st.markdown("### Buy fewer units. Buy bigger ones.")
    st.write(
        f"A $1,000 meter earns about the same return as a $200 meter and sells just as fast. "
        f"Hours limit this business. Money does not."
    )
    a, b = st.columns([1, 1])
    with a:
        st.markdown("**Where the money comes from**")
        st.write(
            f"{D['repeat_share']*100:.0f}% of profit (${D['paid_profit']:,.0f}) came from units I bought "
            f"to resell. The rest (${D['free_profit']:,.0f}) came from things that cost me nothing and "
            f"will not repeat."
        )
        chart("c_bought_vs_free.png")
        chart("c_profit_by_line.png",
              f"Fluke is ${D['fluke_profit']:,.0f} of ${D['profit']:,.0f}, "
              f"or {D['fluke_share']*100:.0f}% of all profit.")
    with b:
        st.markdown("**Bigger meters, same return**")
        st.write(
            "Sorting every Fluke sale by what I paid: the return barely moves across a 10x price "
            "range, but profit per sale rises about 5x. Every sale costs the same search."
        )
        chart("c_ticket_bands.png")
        st.markdown("**Speed, and the one buy that did not have it**")
        chart("c_capital_vs_profit.png",
              f"Average hold time is {D['all_days']:.1f} days. The guitar took 30.")

# ---------------- tab 2 ----------------
with tab2:
    st.markdown("### Should I put more hours and money into one product line?")
    m = st.columns(4)
    m[0].metric("Fluke sales", f"{D['fluke_n']}")
    m[1].metric("Return", f"{D['fluke_roi']*100:.1f}%")
    m[2].metric("Avg hold", f"{D['fluke_days']:.1f} days")
    m[3].metric("Share of profit", f"{D['fluke_share']*100:.0f}%")

    a, b = st.columns([1, 1])
    with a:
        st.markdown("**Fluke's strategy, and what it hands me**")
        st.write(
            "A multimeter is safety equipment. Read a live wire as dead and someone gets hurt. "
            "Fluke built its position on that fact: compete on trust and longevity, never on price. "
            "That is a deliberate strategy. Budget meters fail in a few years. A Fluke runs fifteen "
            "to twenty-five, and electricians still carry meters from the 1990s."
        )
        st.write(
            "A product built to refuse to wear out is why a ten-year-old unit still sells for real "
            "money. Their strategy is my margin: in the used market, a Fluke competes only with "
            "other Flukes."
        )
        st.info(
            "Pushing back on myself: Fluke is not a monopoly. Klein sells credible meters for less "
            "and plenty of electricians carry both. Fluke owns the premium and industrial end, which "
            "is the end that holds value."
        )
        st.markdown("**Demand is not what limits me**")
        chart("c_fluke_demand.png",
              f"{sum(aa.FLUKE_DEMAND.values()):,} recent sales across five models. "
              f"My {D['fluke_n']} sales are about "
              f"{D['fluke_n']/sum(aa.FLUKE_DEMAND.values())*100:.1f}% of that.")
    with b:
        st.markdown("**My hours are what limit me**")
        st.write(
            f"I search about {aa.SEARCH_HOURS_NOW:.0f} hours a week and find a meter every one to "
            f"two weeks. I pass on listings I can see, because I run out of time before I run out "
            f"of meters. {aa.SEARCH_HOURS_MAX:.0f} hours is the most I would spend."
        )
        chart("c_scale_levers.png", "Two levers: hours and ticket. Both together beat either alone.")
        st.warning(
            "The assumption doing the work: more hours means proportionally more meters. I have not "
            "tested it. If good listings are rarer than my time is short, the fifth hour finds less "
            "than the first. Tracking search hours is the first thing I would fix."
        )

# ---------------- tab 3 ----------------
with tab3:
    st.markdown("### The full deck")
    st.caption("Part one is the profitability review. Part two is the Fluke case study.")
    for i in range(1, 12):
        slide(i)

st.divider()
st.caption(
    "Sources: Hançerlioğulları, Şen & Ağca Aktunç (2016), Int. J. Physical Distribution & Logistics "
    "Management. · Tanlamai, Khern-am-nuai & Adulyasak (2024), AI & Society. · Revine (2025). · "
    "SparkShift (2026)."
)
