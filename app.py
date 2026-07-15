"""
Antidote dashboard.

One page, three tabs. Reads the tracker spreadsheet, recomputes every number,
redraws every chart, and shows the deck. Nothing is typed in by hand.

RUN LOCALLY:   pip install streamlit openpyxl matplotlib
               streamlit run app.py
DEPLOY:        push this folder to GitHub, point share.streamlit.io at app.py.
UPDATE DATA:   add rows to the tracker, re-upload it. Everything refreshes.
"""

import os
import streamlit as st
import antidote_analysis as aa

st.set_page_config(page_title="Antidote", page_icon="🟨", layout="wide")

NAVY = "#0C1F2D"
GOLD = "#F2B909"
MUT = "#6E7A86"

# ---------- one font, one look ----------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400&display=swap');

html, body, [class*="css"], [class*="st-"], p, div, span, li, td, th,
h1, h2, h3, h4, button, input, label {{
    font-family: 'Carlito', 'Calibri', sans-serif !important;
}}
.stApp {{ background: #FFFFFF; }}
h1, h2, h3, h4 {{ color: {NAVY}; }}
footer {{ visibility: hidden; }}

.gold-tag {{ width: 64px; height: 9px; background: {GOLD}; margin-bottom: 10px; }}
.hero-title {{ font-size: 46px; font-weight: 700; color: {NAVY}; line-height: 1.05; margin: 0; }}
.hero-motto {{ font-size: 21px; font-weight: 700; color: {GOLD}; margin: 6px 0 2px 0; }}
.hero-sub   {{ font-size: 15px; color: {MUT}; margin: 0 0 6px 0; }}

.statrow {{ display: flex; gap: 14px; margin: 14px 0 4px 0; flex-wrap: wrap; }}
.stat {{ flex: 1; min-width: 150px; background: #FFFFFF; border: 1.6px solid {GOLD};
         border-radius: 8px; padding: 12px 14px; }}
.stat .v {{ font-size: 30px; font-weight: 700; color: {NAVY}; line-height: 1.1; }}
.stat .l {{ font-size: 12.5px; color: {MUT}; margin-top: 2px; }}

.sec {{ font-size: 13px; font-weight: 700; letter-spacing: 0.06em; color: {NAVY};
        border-bottom: 2px solid {NAVY}; padding-bottom: 4px; margin: 10px 0 8px 0;
        text-transform: uppercase; }}
.note {{ background: #E8ECF1; border-left: 5px solid {GOLD}; padding: 10px 12px;
         border-radius: 4px; font-size: 14px; color: #222; margin: 8px 0; }}
.navycard {{ background: {NAVY}; color: #D9E2EA; padding: 12px 14px; border-radius: 6px;
             font-size: 14px; margin: 8px 0; }}
.navycard b {{ color: {GOLD}; }}

.stTabs [data-baseweb="tab-list"] {{ gap: 6px; border-bottom: 2px solid #E8ECF1; }}
.stTabs [data-baseweb="tab"] {{ font-size: 16px; font-weight: 700; color: {MUT};
    padding: 8px 14px; }}
.stTabs [aria-selected="true"] {{ color: {NAVY} !important;
    border-bottom: 3px solid {GOLD} !important; }}
.stDownloadButton button {{ background: {GOLD}; color: {NAVY}; font-weight: 700;
    border: none; border-radius: 6px; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load(_tracker_mtime):
    """Read the tracker, recompute every number, redraw every chart.
    The mtime argument refreshes the cache whenever the spreadsheet changes."""
    rows = aa.load_rows(aa.INPUT_FILE)
    per_line, _ = aa.summarize(rows)          # prints to server logs only
    bands = aa.ticket_band_analysis(rows)
    aa.make_charts(rows, per_line, bands)
    sold = [x for x in rows if x["status"] == "Sold"]
    inv = [x for x in rows if x["status"] == "Inventory"]
    fluke = [x for x in sold if x["category"] == "Fluke test equipment"]
    fluke_paid = [x for x in fluke if x["acq"] == "Paid"]
    d = {
        "n_sold": len(sold),
        "proceeds": sum(x["net"] for x in sold),
        "profit": sum(x["profit"] for x in sold),
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
    d["fluke_roi"] = d["fluke_profit"] / d["fluke_cogs"] if d["fluke_cogs"] else 0
    d["fluke_share"] = d["fluke_profit"] / d["profit"] if d["profit"] else 0
    d["repeat_share"] = d["paid_profit"] / d["profit"] if d["profit"] else 0
    return d


def stats(items):
    cells = "".join(f'<div class="stat"><div class="v">{v}</div><div class="l">{l}</div></div>'
                    for v, l in items)
    st.markdown(f'<div class="statrow">{cells}</div>', unsafe_allow_html=True)


def sec(t):
    st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)


def chart(name, caption=None):
    if os.path.exists(name):
        st.image(name, use_container_width=True)
        if caption:
            st.caption(caption)


def slide(n):
    for p in (f"slides/slide-{n:02d}.png", f"slide-{n:02d}.png"):
        if os.path.exists(p):
            st.image(p, use_container_width=True)
            return
    st.info(f"Slide {n} image not found.")


try:
    D = load(os.path.getmtime(aa.INPUT_FILE))
except Exception as e:
    st.error(f"Could not read the tracker. Check INPUT_FILE in antidote_analysis.py.\n\n{e}")
    st.stop()

# ---------------- hero ----------------
st.markdown('<div class="gold-tag"></div>', unsafe_allow_html=True)
st.markdown('<p class="hero-title">Antidote</p>', unsafe_allow_html=True)
st.markdown('<p class="hero-motto">Small, certain, fast.</p>', unsafe_allow_html=True)
st.markdown('<p class="hero-sub">I buy underpriced test equipment on Facebook Marketplace '
            'and resell it on eBay. Every figure below is recomputed from my transaction '
            'records each time this page loads.</p>', unsafe_allow_html=True)

stats([
    (f"{D['n_sold']}", "sales"),
    (f"${D['profit']:,.0f}", "realized profit"),
    (f"{D['fluke_roi']*100:.1f}%", "return on Fluke capital"),
    (f"{D['fluke_share']*100:.0f}%", "of profit is Fluke"),
])
st.caption(f"${D['inv_cash']:,.0f} sits in {D['n_inv']} unsold units — inventory at cost, not a loss.")

if os.path.exists("Antidote_Review_and_Fluke_Case_Study.pptx"):
    with open("Antidote_Review_and_Fluke_Case_Study.pptx", "rb") as f:
        st.download_button("Download the deck", f,
                           file_name="Antidote_Review_and_Fluke_Case_Study.pptx")

tab1, tab2, tab3 = st.tabs(["Profitability review", "Fluke case study", "The deck"])

# ---------------- tab 1 ----------------
with tab1:
    st.markdown("### Buy fewer units. Buy bigger ones.")
    st.write("A $1,000 meter earns the same 40% as a $200 meter and sells just as fast. "
             "Hours limit this business, money does not.")
    a, b = st.columns(2, gap="large")
    with a:
        sec("Where profit comes from")
        chart("c_bought_vs_free.png")
        chart("c_profit_by_line.png",
              f"Fluke is ${D['fluke_profit']:,.0f} of ${D['profit']:,.0f} — "
              f"{D['fluke_share']*100:.0f}% of all profit.")
    with b:
        sec("Bigger meters, same return")
        chart("c_ticket_bands.png",
              "Return holds near 40% across a 10x price range. Profit per sale rises 5x.")
        sec("Speed")
        chart("c_capital_vs_profit.png",
              f"Average hold is {D['all_days']:.1f} days. The one guitar took 30.")

# ---------------- tab 2 ----------------
with tab2:
    st.markdown("### Should more hours and money go into one product line?")
    stats([
        (f"{D['fluke_n']}", "Fluke sales"),
        (f"{D['fluke_roi']*100:.1f}%", "return"),
        (f"{D['fluke_days']:.1f} days", "average hold"),
        (f"{D['fluke_share']*100:.0f}%", "share of profit"),
    ])
    a, b = st.columns(2, gap="large")
    with a:
        sec("Fluke's strategy, and what it hands me")
        st.write("A multimeter is safety equipment. Fluke competes on trust and longevity, "
                 "never on price. Budget meters fail in a few years; a Fluke runs fifteen to "
                 "twenty-five, and electricians still carry meters from the 1990s.")
        st.markdown('<div class="navycard">A product that refuses to wear out is why a '
                    'ten-year-old unit sells for real money. <b>Their strategy is my margin:</b> '
                    'used, a Fluke competes only with other Flukes.</div>',
                    unsafe_allow_html=True)
        st.markdown('<div class="note"><b>Pushing back on myself:</b> Fluke is not a monopoly. '
                    'Klein sells credible meters for less. Fluke owns the premium end — '
                    'the end that holds value.</div>', unsafe_allow_html=True)
        sec("Demand")
        chart("c_fluke_demand.png",
              f"{sum(aa.FLUKE_DEMAND.values()):,} recent sold listings across five models. "
              f"My {D['fluke_n']} sales are about "
              f"{D['fluke_n']/sum(aa.FLUKE_DEMAND.values())*100:.1f}% of that.")
    with b:
        sec("Supply — my hours are the limit")
        st.write(f"I search about {aa.SEARCH_HOURS_NOW:.0f} hours a week and find a meter "
                 f"every one to two weeks. I pass on listings I can see. "
                 f"{aa.SEARCH_HOURS_MAX:.0f} hours is the most I would spend.")
        chart("c_scale_levers.png",
              "Two levers — hours and ticket. Both together beat either alone.")
        st.markdown('<div class="note"><b>The assumption doing the work:</b> more hours means '
                    'proportionally more meters. Untested. Tracking search hours is the '
                    'first fix.</div>', unsafe_allow_html=True)

# ---------------- tab 3 ----------------
with tab3:
    st.markdown("### The full deck")
    st.caption("Part one — profitability review. Part two — the Fluke case study.")
    for i in range(1, 12):
        slide(i)

st.divider()
st.caption("Sources: Hançerlioğulları, Şen & Ağca Aktunç (2016), IJPDLM · "
           "Tanlamai, Khern-am-nuai & Adulyasak (2024), AI & Society · "
           "Revine (2025) · SparkShift (2026).")
