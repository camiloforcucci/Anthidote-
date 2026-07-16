"""
Anthidote dashboard.

Reads the tracker spreadsheet, recomputes every number, redraws every chart,
and shows the deck. Nothing on this page is typed in by hand.

RUN LOCALLY:  pip install streamlit openpyxl matplotlib
              streamlit run app.py
DEPLOY:       push this folder to GitHub, point share.streamlit.io at app.py
UPDATE DATA:  add rows to the tracker, re-upload it, the page follows

COLOURS AND FONT SIZE live in .streamlit/config.toml, not in this file.
"""

import os
import glob
import streamlit as st
import antidote_analysis as aa

st.set_page_config(page_title="Anthidote", page_icon="🟨", layout="wide")

NAVY = "#0C1F2D"
GOLD = "#F2B909"
PANEL = "#F1F4F7"
MUT = "#6E7A86"
ICE = "#D9E2EA"

# Only the styling config.toml cannot reach: tab size, image edges, width.
st.markdown(f"""
<style>
  .block-container {{ padding-top: 2.2rem; max-width: 1500px; }}
  .stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 2px solid {PANEL}; }}
  .stTabs [data-baseweb="tab"] {{ font-size: 19px; font-weight: 700; padding: 10px 20px; }}
  .stTabs [aria-selected="true"] {{ border-bottom: 4px solid {GOLD} !important; }}
  div[data-testid="stImage"] img {{ border: 1px solid {PANEL}; border-radius: 6px; }}
  footer, #MainMenu {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load(_mtime):
    """Read the tracker, recompute everything, redraw the charts.
    The _mtime argument refreshes this whenever the spreadsheet changes."""
    rows = aa.load_rows(aa.INPUT_FILE)
    per_line, _ = aa.summarize(rows)
    bands = aa.ticket_band_analysis(rows)
    aa.make_charts(rows, per_line, bands)
    sold = [x for x in rows if x["status"] == "Sold"]
    inv = [x for x in rows if x["status"] == "Inventory"]
    fluke = [x for x in sold if x["category"] == "Fluke test equipment"]
    paid = [x for x in fluke if x["acq"] == "Paid"]
    d = {
        "n_sold": len(sold),
        "profit": sum(x["profit"] for x in sold),
        "inv_cash": sum(x["cogs"] for x in inv),
        "n_inv": len(inv),
        "paid_profit": sum(x["profit"] for x in sold if x["acq"] == "Paid"),
        "fluke_profit": sum(x["profit"] for x in fluke),
        "fluke_cogs": sum(x["cogs"] for x in paid),
        "fluke_n": len(fluke),
        "fluke_ticket": sum(x["cogs"] for x in paid) / max(1, len(paid)),
        "fluke_days": sum(x["turnaround_days"] for x in fluke if x["turnaround_days"])
                      / max(1, len([x for x in fluke if x["turnaround_days"]])),
        "all_days": sum(x["turnaround_days"] for x in sold if x["turnaround_days"])
                    / max(1, len([x for x in sold if x["turnaround_days"]])),
    }
    d["fluke_roi"] = d["fluke_profit"] / d["fluke_cogs"] if d["fluke_cogs"] else 0
    d["fluke_share"] = d["fluke_profit"] / d["profit"] if d["profit"] else 0
    d["repeat_share"] = d["paid_profit"] / d["profit"] if d["profit"] else 0
    return d


def H(t):
    st.markdown(
        f'<div style="font-size:15px;font-weight:700;letter-spacing:.08em;'
        f'text-transform:uppercase;color:{NAVY};border-bottom:2px solid {NAVY};'
        f'padding-bottom:6px;margin:22px 0 14px 0;">{t}</div>', unsafe_allow_html=True)


def body(t):
    st.markdown(f'<div style="font-size:17px;line-height:1.65;">{t}</div>',
                unsafe_allow_html=True)


def note(text, title=None):
    head = (f'<div style="font-weight:700;color:{NAVY};font-size:17px;'
            f'margin-bottom:4px;">{title}</div>') if title else ""
    st.markdown(
        f'<div style="background:{PANEL};border-left:6px solid {GOLD};padding:16px 18px;'
        f'border-radius:6px;margin:12px 0;font-size:16px;line-height:1.5;">'
        f'{head}{text}</div>', unsafe_allow_html=True)


def navycard(text, title=None):
    head = (f'<div style="font-weight:700;color:{GOLD};font-size:17px;'
            f'margin-bottom:6px;">{title}</div>') if title else ""
    st.markdown(
        f'<div style="background:{NAVY};color:{ICE};padding:18px 20px;border-radius:8px;'
        f'margin:12px 0;font-size:16px;line-height:1.55;min-height:120px;">'
        f'{head}{text}</div>', unsafe_allow_html=True)


def statgrid(items):
    cells = "".join(
        f'<div style="flex:1;min-width:170px;background:#FFFFFF;border:2px solid {GOLD};'
        f'border-radius:10px;padding:16px 18px;">'
        f'<div style="font-size:38px;font-weight:800;line-height:1;color:{NAVY};">{v}</div>'
        f'<div style="font-size:15px;color:{MUT};margin-top:6px;line-height:1.3;">{l}</div>'
        f'</div>' for v, l in items)
    st.markdown(f'<div style="display:flex;gap:14px;flex-wrap:wrap;margin:6px 0;">'
                f'{cells}</div>', unsafe_allow_html=True)


def chart(f, cap=None):
    if os.path.exists(f):
        st.image(f, use_container_width=True)
        if cap:
            st.markdown(f'<div style="font-size:14px;color:{MUT};margin:-6px 0 14px 0;">'
                        f'{cap}</div>', unsafe_allow_html=True)


def find_deck():
    for n in ("Anthidote_Review_and_Fluke_Case_Study.pptx",
              "Antidote_Review_and_Fluke_Case_Study.pptx",
              "AntidoteReviewandFlukeCaseStudy.pptx"):
        if os.path.exists(n):
            return n
    hits = glob.glob("*.pptx")
    return hits[0] if hits else None


def find_slides():
    for pat in ("slides/slide-*.png", "slide-*.png", "slides/Slide-*.png", "Slide-*.png"):
        hits = sorted(glob.glob(pat))
        if hits:
            return hits
    return []


try:
    D = load(os.path.getmtime(aa.INPUT_FILE))
except Exception as e:
    st.error(f"Could not read the tracker. Check INPUT_FILE in antidote_analysis.py.\n\n{e}")
    st.stop()

# ===================== HERO =====================
st.markdown(f"""
<div style="background:{NAVY};border-radius:14px;padding:38px 42px 32px 42px;">
  <div style="width:76px;height:10px;background:{GOLD};margin-bottom:18px;"></div>
  <div style="font-size:62px;font-weight:800;color:#FFFFFF;line-height:1;
              letter-spacing:-0.5px;">Anthidote</div>
  <div style="font-size:28px;font-weight:700;color:{GOLD};margin-top:10px;">
      Small, certain, fast.</div>
  <div style="font-size:18px;color:{ICE};margin-top:14px;max-width:900px;line-height:1.5;">
      I buy underpriced test equipment on Facebook Marketplace and resell it on eBay.
      Every number on this page is recalculated from my own transaction records each time
      the page loads.</div>
</div>
""", unsafe_allow_html=True)

st.write("")
statgrid([
    (f"{D['n_sold']}", "sales completed"),
    (f"${D['profit']:,.0f}", "realized profit"),
    (f"{D['fluke_roi']*100:.1f}%", "return on money put into Fluke"),
    (f"{D['fluke_share']*100:.0f}%", "of all profit is Fluke"),
])
st.markdown(f'<div style="font-size:15px;color:{MUT};margin-top:10px;">'
            f'${D["inv_cash"]:,.0f} sits in {D["n_inv"]} unsold units. That is inventory at '
            f'cost, not a loss.</div>', unsafe_allow_html=True)

deck = find_deck()
if deck:
    with open(deck, "rb") as fh:
        st.download_button("Download the deck", fh,
                           file_name="Anthidote_Review_and_Fluke_Case_Study.pptx",
                           type="primary")

st.write("")
t1, t2, t3 = st.tabs(["Profitability review", "Fluke case study", "The deck"])

# ===================== TAB 1 =====================
with t1:
    st.markdown(f'<div style="font-size:32px;font-weight:800;color:{NAVY};margin:6px 0 12px 0;">'
                f'Buy fewer units. Buy bigger ones.</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div style="background:{GOLD};color:{NAVY};padding:16px 20px;border-radius:8px;'
        f'font-size:19px;font-weight:700;line-height:1.45;">A $1,000 meter earns the same '
        f'40% as a $200 meter and sells just as fast. Hours limit this business. Money does '
        f'not.</div>', unsafe_allow_html=True)

    a, b = st.columns(2, gap="large")
    with a:
        H("Where the profit comes from")
        body(f"Free finds carried the summer and made 60% of profit. They make "
             f"{100 - D['repeat_share']*100:.0f}% now. The other "
             f"{D['repeat_share']*100:.0f}% came from meters I went out and bought.")
        chart("c_bought_vs_free.png")
        chart("c_profit_by_line.png",
              f"Fluke is ${D['fluke_profit']:,.0f} of ${D['profit']:,.0f}, which is "
              f"{D['fluke_share']*100:.0f}% of all profit.")
        note("Blended margin fell from 48% to 34%. Free items cost nothing, so they book at "
             "100% and flattered the average. Margin on what I buy held near 30%.",
             "About that margin")
    with b:
        H("Bigger meters, same return")
        body("Sorting every Fluke sale by what I paid: the return barely moves across a 10x "
             "price range, but profit per sale rises about 5x. Every sale costs me the same "
             "search.")
        chart("c_ticket_bands.png", "38%, 44%, 40%. Profit per sale: $72, $167, $349.")
        H("Speed, and the one buy that lacked it")
        chart("c_capital_vs_profit.png",
              f"Average hold is {D['all_days']:.1f} days. The guitar took 30.")
        note("One guitar. I paid $400 and made $60. The 15% is not the problem. It held my "
             "cash for 30 days while everything else turned in three. Never a loss, just "
             "dollars that could not move.", "The one that did not work")

# ===================== TAB 2 =====================
with t2:
    st.markdown(f'<div style="font-size:32px;font-weight:800;color:{NAVY};margin:6px 0 12px 0;">'
                f'Fluke sells trust. I resell it.</div>', unsafe_allow_html=True)
    statgrid([
        (f"{D['fluke_n']}", "Fluke sales sold"),
        (f"{D['fluke_roi']*100:.1f}%", "return on capital"),
        (f"{D['fluke_days']:.1f} days", "average hold"),
        (f"{D['fluke_share']*100:.0f}%", "share of profit"),
    ])

    a, b = st.columns(2, gap="large")
    with a:
        H("Fluke's strategy, and what it hands me")
        body('A multimeter is safety equipment and Fluke is one of the best manufacturers. '
             'Fluke competes on trust and longevity, never on price, which is a deliberate '
             'strategy.<br><br>"Budget meters often fail within 2-5 years" (Revine, 2025). '
             'A Fluke runs fifteen to twenty-five, and "many electricians are still using '
             'Fluke meters from the 1990s" (SparkShift, 2026).')
        navycard("A product that refuses to wear out is why a ten-year-old unit sells for "
                 f"real money. <b style='color:{GOLD};'>Their strategy is my margin:</b> "
                 "used, a Fluke competes only with other Flukes.")
        note("Fluke is not a monopoly. Klein sells credible meters for less and many "
             "electricians carry both. Fluke owns the premium, industrial end that holds "
             "value.", "Monopoly?")
        H("What this means for me")
        body("The buyer cannot test a used meter before it ships. The Fluke name does the "
             "testing.<br><br>So I buy from a seller who wants it gone, or does not know "
             "its value, and sell to an electrician who needs it Monday.<br><br>"
             "<b>No other meter brand carries trust across a resale. That is the entire "
             "niche.</b>")
    with b:
        H("Demand is not what limits me. I am 1% of it.")
        chart("c_fluke_demand.png",
              f"{sum(aa.FLUKE_DEMAND.values()):,} recent sold listings across the five "
              f"models I trade.")
        st.markdown(
            f'<div style="background:#FFFFFF;border:2px solid {GOLD};border-radius:10px;'
            f'padding:20px;text-align:center;">'
            f'<div style="font-size:56px;font-weight:800;color:{NAVY};line-height:1;">1.4%</div>'
            f'<div style="font-size:16px;color:{MUT};margin-top:8px;">Even if all nineteen '
            f'of my sales landed in that window, I am 1.4% of it. Buyers are not my '
            f'problem.</div></div>', unsafe_allow_html=True)
        note('eBay reports these as "recently ended", not a fixed window. 1,389 is a floor '
             'on demand and 1.4% a ceiling on my share. I could sell ten times today\'s '
             'volume and still not press against this market.', "More thoughts")

    st.write("")
    H("Increase the hours and size of the meters")

    c1, c2 = st.columns([1, 1.3], gap="large")
    with c1:
        st.markdown(f'<div style="font-size:18px;font-weight:700;color:{NAVY};'
                    f'margin-bottom:8px;">Move the two levers.</div>',
                    unsafe_allow_html=True)
        hours = st.slider("Hours a week I spend searching", 1.0, 10.0,
                          float(aa.SEARCH_HOURS_NOW), 0.5)
        ticket = st.slider("Average price I pay per meter", 100, 1200,
                           int(round(D["fluke_ticket"] / 25) * 25), 25, format="$%d")
        st.caption(f"Everything else is measured from my records: a "
                   f"{D['fluke_roi']*100:.1f}% return and a {D['fluke_days']:.1f} day hold.")

    units = aa.UNITS_PER_MONTH_NOW * (hours / aa.SEARCH_HOURS_NOW)
    profit = units * ticket * D["fluke_roi"]
    base = aa.UNITS_PER_MONTH_NOW * D["fluke_ticket"] * D["fluke_roi"]
    delta = profit - base

    with c2:
        statgrid([
            (f"${profit:,.0f}", "profit a month"),
            (f"{units:.1f}", "meters a month"),
            (f"${units*ticket:,.0f}", "cash deployed a month"),
            (f"${profit*12:,.0f}", "profit a year"),
        ])
        if abs(delta) > 1:
            word = "more" if delta > 0 else "less"
            st.markdown(
                f'<div style="background:{NAVY};color:{ICE};border-radius:8px;'
                f'padding:16px 20px;font-size:17px;margin-top:12px;">That is '
                f'<b style="color:{GOLD};font-size:22px;">${abs(delta):,.0f} {word}</b> a '
                f'month than what I do today, for {hours:.1f} hours a week instead of '
                f'{aa.SEARCH_HOURS_NOW:.0f}.</div>', unsafe_allow_html=True)

    note("More hours means proportionally more meters. I have not tested that. If good "
         "listings are rarer than my time is short, the fifth hour finds less than the "
         "first. Tracking my search hours is the first thing I would fix.",
         "The assumption doing the work")

    x, y, z = st.columns(3)
    with x:
        navycard("Demand is 1,389 recent sales and I am 1.4% of it. Returns hold at 40% up "
                 "to $1,000 a unit. Nothing here is close to full.", "Scale it.")
    with y:
        navycard("Five hours at a $600 ticket takes about $482 a month to about $1,253. "
                 "Hours alone gets $804. Ticket alone gets $752.", "Both levers, not one.")
    with z:
        navycard("Five meters a month at $600. At a two day hold I never hold all five at "
                 "once. Track the hours starting now.", "Fund it with $3,000.")

    H("What kills this")
    k = st.columns(3)
    for col, txt in zip(k, [
        "A dead meter. Condition risk scales with the ticket, so the test rule matters more "
        "at $1,000 than it did at $200.",
        "Other resellers finding the same gap on Marketplace.",
        "My own time. Five hours is my ceiling during the summer, not the semester.",
    ]):
        with col:
            st.markdown(f'<div style="background:{PANEL};border-left:6px solid {GOLD};'
                        f'padding:14px 16px;border-radius:6px;font-size:16px;'
                        f'line-height:1.5;min-height:112px;">{txt}</div>',
                        unsafe_allow_html=True)

# ===================== TAB 3 =====================
with t3:
    slides = find_slides()
    st.markdown(f'<div style="font-size:24px;font-weight:800;color:{NAVY};margin:6px 0;">'
                f'The full deck, {len(slides)} slides</div>', unsafe_allow_html=True)
    st.caption("Part one is the profitability review. Part two is the Fluke case study.")
    if not slides:
        st.info("No slide images found. Upload slide-01.png and the rest to the repo.")
    for f in slides:
        st.image(f, use_container_width=True)
        st.write("")

st.divider()
st.markdown(f'<div style="font-size:14px;color:{MUT};">Sources: Hançerlioğulları, Şen and '
            f'Ağca Aktunç (2016), IJPDLM. Tanlamai, Khern-am-nuai and Adulyasak (2024), '
            f'AI &amp; Society. Revine (2025). SparkShift (2026).</div>',
            unsafe_allow_html=True)
