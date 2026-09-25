# %% [markdown]
# # EA Take-Private: LBO Analysis of the Largest Buyout Ever
#
# **In August 2026 PIF, Silver Lake and Affinity Partners took Electronic Arts private at $210 per share, about $55 billion.
# Could a financial sponsor have paid that price, and what return does it imply?**
#
# This notebook rebuilds the buyout as a private equity deal team would: sources and uses, the actual $18 billion
# debt package tranche by tranche, a five-year operating model with the planned cost savings, a debt schedule with
# cash sweep, sponsor returns in three scenarios, and the maximum price a sponsor could pay at its target IRR.
#
# | Module | Output |
# |---|---|
# | 1. Data | FY2025-FY2026 financials (10-K, earnings release), share count, deal terms, the closing debt package |
# | 2. Sources & uses | Equity purchase, refinancing, fees; debt, balance-sheet cash, sponsor equity |
# | 3. Operating model | Net bookings, adjusted EBITDA with cost savings, cash compensation, capex, taxes |
# | 4. Debt schedule | Six tranches (TLA, USD/EUR TLB, USD/EUR secured notes, unsecured notes), amortisation, cash sweep |
# | 5. Returns | Exit value, sponsor IRR and MoM in Downside / Base / Upside |
# | 6. Ability to pay | Maximum price per share at a 20% IRR, and the exit multiple needed at $210 |
# | 7. Sensitivity | IRR across entry price and exit multiple |
# | 8. Exports | **Excel model with live formulas** and a scenario switch + charts |
#
# > Educational project, not investment advice. Company and deal figures come from EA's public filings and press coverage
# > of the financing; base rates, the TLA spread and the operating scenarios are assumptions, marked as such.

# %%
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "xlsxwriter"], check=False)

# %%
import os, warnings, datetime as dt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

warnings.filterwarnings("ignore")
pd.set_option("display.float_format", lambda x: f"{x:,.1f}")

# %% [markdown]
# ## 0. Configuration
# All figures in **$ millions** unless stated. In the Excel model these are the blue input cells.

# %%
AUTHOR = "Alessandro Radice"
OUTPUT_DIR = "ea_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)
SCENARIOS = ["Downside", "Base", "Upside"]
ACTIVE = 2                     # scenario shown in Excel: 1 = Downside, 2 = Base, 3 = Upside
YEARS = [2027, 2028, 2029, 2030, 2031]   # five full fiscal years (ending March); the close was August 2026

# %% [markdown]
# ## 1. Data
# FY2026 = year ended 31 March 2026 (earnings release and 10-K). Adjusted EBITDA is on a **net bookings basis**:
# operating income + change in deferred net revenue (bookings − revenue) + D&A + stock-based compensation. On FY2026
# this gives $2,636m, and $18.0bn of debt / $2,636m = 6.8x, the gross leverage reported for the financing.

# %%
FIN = pd.DataFrame({
    "FY2025": dict(net_revenue=7463, net_bookings=7355, op_income=1520, da=356, sbc=642, capex=221, ocf=2079),
    "FY2026": dict(net_revenue=7531, net_bookings=8026, op_income=1162, da=323, sbc=656, capex=230, ocf=2553),
})
FIN.loc["adj_ebitda"] = FIN.loc["op_income"] + (FIN.loc["net_bookings"] - FIN.loc["net_revenue"]) + FIN.loc["da"] + FIN.loc["sbc"]
FIN.loc["margin_on_bookings"] = FIN.loc["adj_ebitda"] / FIN.loc["net_bookings"]

DEAL = dict(
    price=210.0, premium=0.25, shares=250.751484,        # shares in millions (10-K cover, 6 May 2026)
    cash=2980.0, old_notes=1485.0, min_cash=1000.0,       # 31 March 2026
    fin_fee_pct=0.015, oid_pct=0.015, advisory_pct=0.005, # fees: % of debt, OID on term loan B, % of equity value
    marketing_ebitda=3400.0,                               # EBITDA marketed to lenders (incl. ~$400m savings, ~$300m addbacks)
)
DEAL["unaffected"] = DEAL["price"] / (1 + DEAL["premium"])

EURUSD = 2000.0 / 1725.0   # implied by the $2.0bn equivalent of the €1,725m term loan
RATES = dict(sofr=0.0375, euribor=0.0225)   # assumptions
# name, amount $m, floating?, spread or coupon, base ("sofr"/"euribor"/None), annual amortisation % of original, prepayable
DEBT = pd.DataFrame([
    ("Term Loan A",                      3250.0,           True,  0.0175, "sofr",    0.05, True),
    ("Term Loan B (USD)",                6125.0,           True,  0.0350, "sofr",    0.01, True),
    ("Term Loan B (EUR €1,725m)",        1725.0 * EURUSD,  True,  0.0350, "euribor", 0.01, True),
    ("Senior secured notes 2033 (USD)",  2875.0,           False, 0.0725, None,      0.00, False),
    ("Senior secured notes 2033 (EUR €1,080m)", 1080.0 * EURUSD, False, 0.0625, None, 0.00, False),
    ("Senior unsecured notes 2034",      2500.0,           False, 0.0875, None,      0.00, False),
], columns=["tranche", "amount", "floating", "rate", "base", "amort", "prepay"])

OPS = pd.DataFrame({   # by scenario
    "growth":        [0.02, 0.05, 0.08],     # net bookings growth per year
    "savings":       [200.0, 400.0, 700.0],  # run-rate cost savings, 50% in FY2027, full from FY2028
    "exit_multiple": [14.0, 16.0, 19.0],     # EV / adj. EBITDA at exit
}, index=SCENARIOS)
OPS_COMMON = dict(da_pct=FIN.loc["da", "FY2026"] / FIN.loc["net_bookings", "FY2026"],
                  capex_pct=FIN.loc["capex", "FY2026"] / FIN.loc["net_bookings", "FY2026"],
                  sbc_pct=FIN.loc["sbc", "FY2026"] / FIN.loc["net_bookings", "FY2026"],
                  cash_comp_share=0.50,      # share of former stock-based pay now paid in cash
                  tax=0.21, sweep=1.00, target_irr=0.20)
print(FIN.round(3))

# %% [markdown]
# ## 2. Sources & uses

# %%
def sources_uses(price=DEAL["price"]):
    d = DEAL
    debt = DEBT.amount.sum()
    tlb = DEBT.loc[DEBT.tranche.str.startswith("Term Loan B"), "amount"].sum()
    eq_val = d["shares"] * price
    uses = {"Purchase of equity": eq_val, "Refinance existing senior notes": d["old_notes"],
            "Financing fees": debt * d["fin_fee_pct"], "OID on term loan B": tlb * d["oid_pct"],
            "Advisory and other costs": eq_val * d["advisory_pct"]}
    total = sum(uses.values())
    src = {"New debt": debt, "EA balance-sheet cash": d["cash"] - d["min_cash"]}
    src["Sponsor equity"] = total - sum(src.values())
    return uses, src, total

USES, SOURCES, TOTAL = sources_uses()
print(pd.Series(USES).round(1), "\n", pd.Series(SOURCES).round(1), "\nTotal:", round(TOTAL, 1))
EV_ENTRY = DEAL["shares"] * DEAL["price"] + DEAL["old_notes"] - DEAL["cash"]
print(f"Entry EV (equity + debt − cash): ${EV_ENTRY:,.0f}m = {EV_ENTRY / FIN.loc['adj_ebitda', 'FY2026']:.1f}x FY2026 adj. EBITDA")

# %% [markdown]
# ## 3-6. Operating model, debt schedule, returns and ability to pay
# Interest is charged on opening balances (no circularity). Mandatory amortisation first, then 100% of the remaining
# cash above the minimum sweeps the prepayable loans in order (TLA, then the term loan Bs). Notes are bullets.

# %%
def rate_of(row):
    if not row.floating: return row.rate
    return RATES[row.base] + row.rate

def run(scn):
    o = OPS.loc[scn]; c = OPS_COMMON
    bal = DEBT.amount.values.copy(); orig = bal.copy()
    rates = np.array([rate_of(r) for r in DEBT.itertuples()])
    cash = DEAL["min_cash"]; book = FIN.loc["net_bookings", "FY2026"]
    m0 = FIN.loc["margin_on_bookings", "FY2026"]
    rows = []
    for i, y in enumerate(YEARS):
        book = book * (1 + o.growth)
        sav = o.savings * (0.5 if i == 0 else 1.0)
        ebitda = book * m0 + sav
        cash_comp = book * c["sbc_pct"] * c["cash_comp_share"]
        da = book * c["da_pct"]; capex = book * c["capex_pct"]
        interest = float((bal * rates).sum())
        ebt = ebitda - cash_comp - da - interest
        tax = max(ebt, 0) * c["tax"]
        fcf = ebitda - cash_comp - capex - tax - interest
        mand = np.minimum(orig * DEBT.amort.values, bal)
        bal = bal - mand
        avail = max(cash + fcf - mand.sum() - DEAL["min_cash"], 0) * c["sweep"]
        sweep = np.zeros_like(bal)
        for j in np.where(DEBT.prepay.values)[0]:
            s = min(avail, bal[j]); sweep[j] = s; avail -= s
        bal = bal - sweep
        cash = cash + fcf - mand.sum() - sweep.sum()
        rows.append(dict(year=y, bookings=book, savings=sav, ebitda=ebitda, cash_comp=cash_comp, da=da, capex=capex,
                         interest=interest, tax=tax, fcf=fcf, mandatory=mand.sum(), sweep=sweep.sum(),
                         debt=bal.sum(), cash=cash, net_debt=bal.sum() - cash,
                         leverage=(bal.sum() - cash) / ebitda, coverage=ebitda / interest))
    p = pd.DataFrame(rows).set_index("year")
    exit_ev = p.ebitda.iloc[-1] * o.exit_multiple
    exit_eq = exit_ev - p.net_debt.iloc[-1]
    eq_in = SOURCES["Sponsor equity"]
    n = len(YEARS)
    moic = exit_eq / eq_in; irr = moic ** (1 / n) - 1
    # ability to pay: equity that earns the target IRR, back to a price per share
    eq_max = exit_eq / (1 + c["target_irr"]) ** n
    fixed = DEAL["old_notes"] + DEBT.amount.sum() * DEAL["fin_fee_pct"] + \
            DEBT.loc[DEBT.tranche.str.startswith("Term Loan B"), "amount"].sum() * DEAL["oid_pct"] - \
            DEBT.amount.sum() - (DEAL["cash"] - DEAL["min_cash"])
    price_max = (eq_max - fixed) / (DEAL["shares"] * (1 + DEAL["advisory_pct"]))
    # exit multiple needed for the target IRR at $210
    mult_needed = (eq_in * (1 + c["target_irr"]) ** n + p.net_debt.iloc[-1]) / p.ebitda.iloc[-1]
    return p, dict(exit_ebitda=p.ebitda.iloc[-1], exit_ev=exit_ev, exit_net_debt=p.net_debt.iloc[-1], exit_equity=exit_eq,
                   equity_in=eq_in, moic=moic, irr=irr, price_max=price_max, mult_needed=mult_needed,
                   debt_repaid=DEBT.amount.sum() - p.debt.iloc[-1])

RES = {s: run(s) for s in SCENARIOS}
SUM = pd.DataFrame({s: RES[s][1] for s in SCENARIOS})
print(SUM.round(3))
print(RES["Base"][0].round(1).T)

# %% [markdown]
# ## 7. Sensitivity: IRR by entry price and exit multiple (Base operating case)
# With the debt package fixed at $18bn, the entry price changes only the sponsor equity cheque, so the grid is exact.

# %%
SENS_PRICE = [150, 168, 180, 195, 210, 225]
SENS_MULT = [12, 14, 16, 18, 20, 22]
pB = RES["Base"][0]
def irr_at(price, mult):
    _, src, _ = sources_uses(price)
    exit_eq = pB.ebitda.iloc[-1] * mult - pB.net_debt.iloc[-1]
    return (exit_eq / src["Sponsor equity"]) ** (1 / len(YEARS)) - 1
SENS = pd.DataFrame([[irr_at(p, m) for m in SENS_MULT] for p in SENS_PRICE],
                    index=[f"${p}" for p in SENS_PRICE], columns=[f"{m}x" for m in SENS_MULT])
print((SENS * 100).round(1))

# %% [markdown]
# ## 8. Charts

# %%
BG, INK, MUT = "#F7F6F2", "#1A1A18", "#66655E"
plt.rcParams.update({"font.family": "DejaVu Sans", "axes.edgecolor": INK, "xtick.color": MUT, "ytick.color": MUT,
                     "figure.facecolor": BG, "axes.facecolor": BG})
META = {"Author": AUTHOR, "Software": AUTHOR}
fig, ax = plt.subplots(figsize=(10, 5.2))
yrs = [f"FY{str(y)[2:]}" for y in YEARS]
ax.bar(yrs, pB.debt / 1000, color="#1A1A18", width=0.55, label="Gross debt, $bn")
ax2 = ax.twinx()
ax2.plot(yrs, pB.leverage, color="#8A8980", lw=2.5, marker="o", label="Net debt / adj. EBITDA")
for i, v in enumerate(pB.leverage): ax2.text(i, v + 0.15, f"{v:.1f}x", ha="center", fontsize=10, color=INK)
ax.set_ylim(0, 20); ax2.set_ylim(0, 7.5)
ax.spines[["top"]].set_visible(False); ax2.spines[["top"]].set_visible(False)
ax.set_title("Base case: deleveraging from the $18bn package", loc="left", fontsize=13, color=INK)
fig.legend(frameon=False, loc="upper right", bbox_to_anchor=(0.9, 0.88))
plt.tight_layout(); plt.savefig(f"{OUTPUT_DIR}/deleveraging.png", dpi=160, metadata={**META, "Title": "EA deleveraging"}); plt.show()

# %% [markdown]
# ## 9. Excel model with live formulas

# %%
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

XLSX = f"{OUTPUT_DIR}/EA_LBO_Model.xlsx"
wb = xlsxwriter.Workbook(XLSX)
wb.set_properties({"title": "EA Take-Private: LBO Analysis", "author": AUTHOR, "company": AUTHOR,
                   "subject": "Leveraged buyout of Electronic Arts: sources and uses, debt schedule, returns, ability to pay",
                   "keywords": "LBO, private equity, leveraged finance, Electronic Arts", "created": dt.datetime.now(),
                   "comments": "Educational project. Figures from public filings and press coverage; assumptions marked."})
F = dict(font_name="Arial", font_size=10); fmt = lambda **k: wb.add_format({**F, **k})
NUM = '#,##0;(#,##0);"–"'
f_title = fmt(bold=True, font_size=14); f_sub = fmt(italic=True, font_color="#66655E")
f_hdr = fmt(bold=True, font_color="white", bg_color="#1A1A18", align="center"); f_hdrl = fmt(bold=True, font_color="white", bg_color="#1A1A18")
f_b = fmt(bold=True); f_note = fmt(italic=True, font_color="#66655E", text_wrap=True, valign="top")
f_in = fmt(font_color="#0000FF", num_format=NUM); f_in1 = fmt(font_color="#0000FF", num_format="#,##0.0")
f_inp = fmt(font_color="#0000FF", num_format="0.00%"); f_inx = fmt(font_color="#0000FF", num_format='0.0"x"')
f_ind = fmt(font_color="#0000FF", num_format='$#,##0.00'); f_inb = fmt(font_color="#0000FF", align="center")
f_n = fmt(num_format=NUM); f_p = fmt(num_format="0.0%"); f_x = fmt(num_format='0.0"x"'); f_d = fmt(num_format='$#,##0.00')
f_tot = fmt(num_format=NUM, bold=True, top=1); f_totx = fmt(num_format='0.00"x"', bold=True, top=1, bg_color="#ECEBE5")
f_totp = fmt(num_format="0.0%", bold=True, bg_color="#ECEBE5"); f_totd = fmt(num_format='$#,##0.00', bold=True, bg_color="#ECEBE5")
f_lk = fmt(font_color="#008000", num_format=NUM); f_lkp = fmt(font_color="#008000", num_format="0.0%")
f_lkx = fmt(font_color="#008000", num_format='0.0"x"')
f_sel = fmt(font_color="#0000FF", bold=True, bg_color="#FFF2CC", border=1, align="center")
f_chk = fmt(num_format='0.000;-0.000;"OK"', bold=True, font_color="#008000")
f_heat = fmt(num_format="0.0%", align="center", border=1, border_color="#DAD9D2")

def sheet(name, title, sub="$ millions unless stated · FY = fiscal year ending 31 March"):
    ws = wb.add_worksheet(name); ws.hide_gridlines(2); ws.set_column("A:A", 2); ws.set_column("B:B", 46)
    ws.write("B2", title, f_title); ws.write("B3", sub, f_sub); return ws

# Cover
wc = sheet("Cover", "EA Take-Private: LBO Analysis", f"Prepared by {AUTHOR} · Educational project, not investment advice")
wc.set_column("C:C", 96)
for i, (a, b) in enumerate([
    ("Question", "Could a financial sponsor pay $210 per share for EA, and what return does the price imply?"),
    ("Inputs", "Financials, deal terms, the six debt tranches, base rates, scenario drivers and the scenario switch (C5)"),
    ("Sources & Uses", "Equity purchase, refinancing and fees vs new debt, balance-sheet cash and sponsor equity"),
    ("Operating Model", "Net bookings, adjusted EBITDA with cost savings, cash compensation, capex, taxes, free cash flow"),
    ("Debt Schedule", "Tranche by tranche: interest, mandatory amortisation, cash sweep, leverage and coverage"),
    ("Returns", "Exit value, sponsor IRR and MoM, maximum price at the target IRR, exit multiple needed at $210"),
    ("Sensitivity", "IRR by entry price and exit multiple, for the active scenario"),
    ("Colour code", "Blue = hard-coded input · Black = formula · Green = link to another sheet"),
]):
    wc.write(5 + i, 1, a, f_b); wc.write(5 + i, 2, b)

# Inputs
wi = sheet("Inputs", "Inputs"); wi.set_column("C:H", 13); wi.set_column("I:I", 60)
wi.write("B5", "Active scenario (1 = Downside, 2 = Base, 3 = Upside)", f_b); wi.write("C5", ACTIVE, f_sel)
wi.write_formula("D5", '=CHOOSE($C$5,"Downside","Base","Upside")', f_b)
SEL = "Inputs!$C$5"
R = {}
def inp(r, lab, v, f_, note="", key=None, col=2):
    wi.write(r, 1, lab); wi.write(r, col, v, f_)
    if note: wi.write(r, 8, note, f_note)
    if key: R[key] = f"Inputs!${xl_col_to_name(col)}${r + 1}"
r = 7
wi.write(r, 1, "Historical financials", f_hdrl); wi.write(r, 2, "FY2025", f_hdr); wi.write(r, 3, "FY2026", f_hdr); r += 1
HR = {}
for key, lab in [("net_revenue", "Net revenue"), ("net_bookings", "Net bookings"), ("op_income", "Operating income (GAAP)"),
                 ("da", "Depreciation & amortisation"), ("sbc", "Stock-based compensation"), ("capex", "Capital expenditure")]:
    wi.write(r, 1, lab); wi.write(r, 2, float(FIN.loc[key, "FY2025"]), f_in); wi.write(r, 3, float(FIN.loc[key, "FY2026"]), f_in)
    HR[key] = r + 1; r += 1
wi.write(r, 1, "Adj. EBITDA (bookings basis)", f_b)
for cc in "CD":
    wi.write_formula(r, "CD".index(cc) + 2, f"={cc}{HR['op_income']}+({cc}{HR['net_bookings']}-{cc}{HR['net_revenue']})+{cc}{HR['da']}+{cc}{HR['sbc']}", f_tot)
wi.write(r, 8, "Operating income + change in deferred net revenue + D&A + SBC", f_note)
HR["adj_ebitda"] = r + 1; r += 1
wi.write(r, 1, "Adj. EBITDA margin on bookings")
for cc in "CD": wi.write_formula(r, "CD".index(cc) + 2, f"={cc}{HR['adj_ebitda']}/{cc}{HR['net_bookings']}", f_p)
HR["margin"] = r + 1; r += 2
A = lambda k, c="D": f"Inputs!${c}${HR[k]}"
wi.write(r, 1, "Deal terms and balance sheet", f_hdrl); r += 1
for key, lab, v, f_, note in [
    ("price", "Offer price per share ($)", DEAL["price"], f_ind, "All-cash merger consideration"),
    ("premium", "Premium to unaffected price", DEAL["premium"], f_inp, "Reported 25%"),
    ("shares", "Shares outstanding (m)", DEAL["shares"], f_in1, "10-K cover, 6 May 2026"),
    ("cash", "Cash and short-term investments, 31 Mar 2026", DEAL["cash"], f_in, "FY2026 earnings release"),
    ("old_notes", "Existing senior notes, refinanced", DEAL["old_notes"], f_in, "Tendered or defeased at closing"),
    ("min_cash", "Minimum cash kept in the business", DEAL["min_cash"], f_in, "Assumption"),
    ("fin_fee_pct", "Financing fees, % of new debt", DEAL["fin_fee_pct"], f_inp, "Assumption"),
    ("oid_pct", "OID on term loan B", DEAL["oid_pct"], f_inp, "Issued at 98.5"),
    ("advisory_pct", "Advisory and other costs, % of equity value", DEAL["advisory_pct"], f_inp, "Assumption"),
    ("marketing_ebitda", "EBITDA marketed to lenders", DEAL["marketing_ebitda"], f_in, "Incl. ~$400m savings and ~$300m addbacks"),
    ("sofr", "SOFR", RATES["sofr"], f_inp, "Assumption"), ("euribor", "Euribor", RATES["euribor"], f_inp, "Assumption"),
]:
    inp(r, lab, v, f_, note, key); r += 1
wi.write(r, 1, "Unaffected price ($)"); wi.write_formula(r, 2, f"={R['price']}/(1+{R['premium']})", f_d); R["unaffected"] = f"Inputs!$C${r + 1}"; r += 2
wi.write(r, 1, "Debt package at closing (4 Aug 2026)", f_hdrl)
for j, h in enumerate(["Amount", "Floating?", "Spread/coupon", "All-in rate", "Amort. %/yr", "Prepayable?"]): wi.write(r, 2 + j, h, f_hdr)
r += 1; DR0 = r + 1
for t in DEBT.itertuples():
    wi.write(r, 1, t.tranche); wi.write(r, 2, t.amount, f_in); wi.write(r, 3, "Yes" if t.floating else "No", f_inb)
    wi.write(r, 4, t.rate, f_inp)
    base = R["sofr"] if t.base == "sofr" else R["euribor"] if t.base == "euribor" else "0"
    wi.write_formula(r, 5, f'=IF(D{r + 1}="Yes",{base},0)+E{r + 1}' if base != "0" else f'=E{r + 1}', f_p)
    wi.write(r, 6, t.amort, f_inp); wi.write(r, 7, "Yes" if t.prepay else "No", f_inb)
    wi.write(r, 8, {"Term Loan A": "Spread assumed (not disclosed); 5% amortisation assumed"}.get(t.tranche, "Closing 8-K / syndication press"), f_note)
    r += 1
DRn = r
wi.write(r, 1, "Total new debt", f_b); wi.write_formula(r, 2, f"=SUM(C{DR0}:C{DRn})", f_tot); R["debt"] = f"Inputs!$C${r + 1}"; r += 1
wi.write(r, 1, "Term loan B total"); wi.write_formula(r, 2, f"=C{DR0 + 1}+C{DR0 + 2}", f_n); R["tlb"] = f"Inputs!$C${r + 1}"; r += 2
wi.write(r, 1, "Scenario drivers", f_hdrl); [wi.write(r, 2 + j, h, f_hdr) for j, h in enumerate(SCENARIOS + ["► Live"])]; r += 1
for key, lab, f_ in [("growth", "Net bookings growth per year", f_inp), ("savings", "Run-rate cost savings (50% in FY2027)", f_in),
                     ("exit_multiple", "Exit EV / adj. EBITDA", f_inx)]:
    wi.write(r, 1, lab)
    for j, s in enumerate(SCENARIOS): wi.write(r, 2 + j, float(OPS.loc[s, key]), f_)
    wi.write_formula(r, 5, f"=CHOOSE({SEL},C{r + 1},D{r + 1},E{r + 1})", fmt(bold=True, bg_color="#ECEBE5", num_format={"growth": "0.0%", "savings": NUM, "exit_multiple": '0.0"x"'}[key]))
    R[key] = f"Inputs!$F${r + 1}"; r += 1
r += 1
wi.write(r, 1, "Common assumptions", f_hdrl); r += 1
for key, lab, note in [("da_pct", "D&A % of bookings", "FY2026 ratio"), ("capex_pct", "Capex % of bookings", "FY2026 ratio"),
                       ("sbc_pct", "Stock-based compensation % of bookings", "FY2026 ratio"),
                       ("cash_comp_share", "Share of former SBC now paid in cash", "Assumption: unvested RSUs became cash awards"),
                       ("tax", "Cash tax rate", "Assumption"), ("sweep", "Excess cash sweep", "Applied to TLA, then term loan B"),
                       ("target_irr", "Sponsor target IRR", "Typical large-cap buyout hurdle")]:
    inp(r, lab, float(OPS_COMMON[key]), f_inp, note, key); r += 1

# Sources & Uses
wsu = sheet("Sources & Uses", "Sources & uses at closing"); wsu.set_column("C:D", 14)
wsu.write(5, 1, "Uses", f_hdrl); wsu.write(5, 2, "$m", f_hdr); wsu.write(5, 3, "% total", f_hdr)
UL = [("Purchase of equity", f"={R['shares']}*{R['price']}"), ("Refinance existing senior notes", f"={R['old_notes']}"),
      ("Financing fees", f"={R['debt']}*{R['fin_fee_pct']}"), ("OID on term loan B", f"={R['tlb']}*{R['oid_pct']}"),
      ("Advisory and other costs", f"=C7*{R['advisory_pct']}")]
for k, (lab, form) in enumerate(UL):
    wsu.write(6 + k, 1, lab); wsu.write_formula(6 + k, 2, form, f_lk if k == 1 else f_n); wsu.write_formula(6 + k, 3, f"=C{7 + k}/$C$12", f_p)
wsu.write(11, 1, "Total uses", f_b); wsu.write_formula(11, 2, "=SUM(C7:C11)", f_tot)
wsu.write(13, 1, "Sources", f_hdrl); wsu.write(13, 2, "$m", f_hdr); wsu.write(13, 3, "% total", f_hdr)
SL = [("New debt", f"={R['debt']}"), ("EA balance-sheet cash", f"={R['cash']}-{R['min_cash']}"), ("Sponsor equity (residual)", "=C12-C15-C16")]
for k, (lab, form) in enumerate(SL):
    wsu.write(14 + k, 1, lab, f_b if k == 2 else None); wsu.write_formula(14 + k, 2, form, f_tot if k == 2 else f_lk); wsu.write_formula(14 + k, 3, f"=C{15 + k}/$C$18", f_p)
wsu.write(17, 1, "Total sources", f_b); wsu.write_formula(17, 2, "=SUM(C15:C17)", f_tot)
wsu.write(18, 1, "Check: sources − uses", f_b); wsu.write_formula(18, 2, "=ROUND(C18-C12,6)", f_chk)
wsu.write(20, 1, "Entry valuation", f_hdrl); wsu.write(20, 2, "", f_hdr); wsu.write(20, 3, "", f_hdr)
for k, (lab, form, f_) in enumerate([
    ("Entry EV (equity + old notes − cash)", f"=C7+{R['old_notes']}-{R['cash']}", f_n),
    ("EV / FY2026 adj. EBITDA", f"=C22/{A('adj_ebitda')}", f_x),
    ("EV / marketing EBITDA", f"=C22/{R['marketing_ebitda']}", f_x),
    ("Gross leverage: new debt / FY2026 adj. EBITDA", f"={R['debt']}/{A('adj_ebitda')}", f_x),
    ("Gross leverage on marketing EBITDA", f"={R['debt']}/{R['marketing_ebitda']}", f_x),
    ("Equity cushion (sponsor equity / total sources)", "=C17/C18", f_p),
    ("Unaffected price ($)", f"={R['unaffected']}", fmt(font_color="#008000", num_format='$#,##0.00'))]):
    wsu.write(21 + k, 1, lab); wsu.write_formula(21 + k, 2, form, f_)
EQ_IN = "'Sources & Uses'!$C$17"

# Operating model
DS_INT_ROW = 6 + 7 * len(DEBT) + 5 + 1     # Excel row of total cash interest on the Debt Schedule
wo = sheet("Operating Model", "Operating model (active scenario)"); wo.set_column("C:H", 12)
wo.write_formula("B4", '="Scenario: "&Inputs!$D$5', f_b)
cols = ["FY2026A"] + [f"FY{y}E" for y in YEARS]
for j, h in enumerate(cols): wo.write(5, 2 + j, h, f_hdr)
OL = ["Net bookings", "Growth", "Adj. EBITDA before savings", "Cost savings", "Adj. EBITDA", "Margin on bookings",
      "(–) Cash compensation replacing SBC", "(–) D&A", "(–) Cash interest", "EBT (tax basis)", "(–) Cash taxes",
      "(–) Capex", "Free cash flow after interest"]
OR = {k: 6 + i for i, k in enumerate(OL)}
for k in OL: wo.write(OR[k], 1, k, f_b if k in ("Net bookings", "Adj. EBITDA", "Free cash flow after interest") else None)
O = lambda k, C: f"{C}{OR[k] + 1}"
wo.write_formula(OR["Net bookings"], 2, f"={A('net_bookings')}", f_lk)
wo.write_formula(OR["Adj. EBITDA"], 2, f"={A('adj_ebitda')}", f_lk)
wo.write_formula(OR["Margin on bookings"], 2, f"=C{OR['Adj. EBITDA'] + 1}/C{OR['Net bookings'] + 1}", f_p)
for i in range(len(YEARS)):
    C = xl_col_to_name(3 + i); P = xl_col_to_name(2 + i)
    wo.write_formula(OR["Net bookings"], 3 + i, f"={P}{OR['Net bookings'] + 1}*(1+{R['growth']})", f_n)
    wo.write_formula(OR["Growth"], 3 + i, f"={O('Net bookings', C)}/{O('Net bookings', P)}-1", f_p)
    wo.write_formula(OR["Adj. EBITDA before savings"], 3 + i, f"={O('Net bookings', C)}*{A('margin')}", f_n)
    wo.write_formula(OR["Cost savings"], 3 + i, f"={R['savings']}*{0.5 if i == 0 else 1}", f_n)
    wo.write_formula(OR["Adj. EBITDA"], 3 + i, f"={O('Adj. EBITDA before savings', C)}+{O('Cost savings', C)}", f_tot)
    wo.write_formula(OR["Margin on bookings"], 3 + i, f"={O('Adj. EBITDA', C)}/{O('Net bookings', C)}", f_p)
    wo.write_formula(OR["(–) Cash compensation replacing SBC"], 3 + i, f"=-{O('Net bookings', C)}*{R['sbc_pct']}*{R['cash_comp_share']}", f_n)
    wo.write_formula(OR["(–) D&A"], 3 + i, f"=-{O('Net bookings', C)}*{R['da_pct']}", f_n)
    wo.write_formula(OR["(–) Cash interest"], 3 + i, f"=-'Debt Schedule'!{C}${DS_INT_ROW}", f_lk)
    wo.write_formula(OR["EBT (tax basis)"], 3 + i, f"={O('Adj. EBITDA', C)}+{O('(–) Cash compensation replacing SBC', C)}+{O('(–) D&A', C)}+{O('(–) Cash interest', C)}", f_n)
    wo.write_formula(OR["(–) Cash taxes"], 3 + i, f"=-MAX({O('EBT (tax basis)', C)},0)*{R['tax']}", f_n)
    wo.write_formula(OR["(–) Capex"], 3 + i, f"=-{O('Net bookings', C)}*{R['capex_pct']}", f_n)
    wo.write_formula(OR["Free cash flow after interest"], 3 + i, f"={O('Adj. EBITDA', C)}+{O('(–) Cash compensation replacing SBC', C)}+{O('(–) Cash interest', C)}+{O('(–) Cash taxes', C)}+{O('(–) Capex', C)}", f_tot)

# Debt schedule
wds = sheet("Debt Schedule", "Debt schedule (active scenario)"); wds.set_column("C:H", 12)
for j, h in enumerate(cols): wds.write(5, 2 + j, "Closing" if j == 0 else h, f_hdr)
r = 6; BAL = {}; INT = {}; MAND = {}; SWP = {}
nT = len(DEBT)
for k, t in enumerate(DEBT.itertuples()):
    ir = DR0 + k   # Inputs row of this tranche
    wds.write(r, 1, t.tranche, f_b)
    rows = dict(open=r + 1, mand=r + 2, sweep=r + 3, close=r + 4, intr=r + 5)
    for lab, off in [("  Opening balance", 1), ("  Mandatory amortisation", 2), ("  Cash sweep", 3), ("  Closing balance", 4), ("  Interest (on opening)", 5)]:
        wds.write(r + off, 1, lab)
    wds.write_formula(rows["close"], 2, f"=Inputs!$C${ir}", f_lk)
    for i in range(len(YEARS)):
        C = xl_col_to_name(3 + i); P = xl_col_to_name(2 + i)
        wds.write_formula(rows["open"], 3 + i, f"={P}{rows['close'] + 1}", f_n)
        wds.write_formula(rows["mand"], 3 + i, f"=-MIN(Inputs!$C${ir}*Inputs!$G${ir},{C}{rows['open'] + 1})", f_n)
        wds.write_formula(rows["close"], 3 + i, f"={C}{rows['open'] + 1}+{C}{rows['mand'] + 1}+{C}{rows['sweep'] + 1}", f_tot)
        wds.write_formula(rows["intr"], 3 + i, f"={C}{rows['open'] + 1}*Inputs!$F${ir}", f_n)
    BAL[k] = rows; r += 7
# totals and sweep waterfall
TR = dict(mand=r, avail=r + 1, cash_open=r + 2, cash_close=r + 3, debt=r + 4, int=r + 5, netdebt=r + 6, lev=r + 7, cov=r + 8)
for lab, key in [("Total mandatory amortisation", "mand"), ("Cash available for sweep", "avail"), ("Opening cash", "cash_open"),
                 ("Closing cash", "cash_close"), ("Total gross debt", "debt"), ("Total cash interest", "int"),
                 ("Net debt", "netdebt"), ("Net debt / adj. EBITDA", "lev"), ("Adj. EBITDA / cash interest", "cov")]:
    wds.write(TR[key], 1, lab, f_b)
wds.write_formula(TR["debt"], 2, "=" + "+".join(f"C{BAL[k]['close'] + 1}" for k in range(nT)), f_tot)
wds.write_formula(TR["cash_close"], 2, f"={R['min_cash']}", f_lk)
wds.write_formula(TR["netdebt"], 2, f"=C{TR['debt'] + 1}-C{TR['cash_close'] + 1}", f_n)
wds.write_formula(TR["lev"], 2, f"=C{TR['netdebt'] + 1}/{A('adj_ebitda')}", f_x)
prepay_order = [k for k, t in enumerate(DEBT.itertuples()) if t.prepay]
for i in range(len(YEARS)):
    C = xl_col_to_name(3 + i); P = xl_col_to_name(2 + i)
    wds.write_formula(TR["mand"], 3 + i, "=" + "+".join(f"{C}{BAL[k]['mand'] + 1}" for k in range(nT)), f_n)
    wds.write_formula(TR["cash_open"], 3 + i, f"={P}{TR['cash_close'] + 1}", f_n)
    wds.write_formula(TR["avail"], 3 + i, f"=MAX({C}{TR['cash_open'] + 1}+'Operating Model'!{C}{OR['Free cash flow after interest'] + 1}+{C}{TR['mand'] + 1}-{R['min_cash']},0)*{R['sweep']}", f_n)
    prev = []
    for k in range(nT):
        if k in prepay_order:
            left = f"{C}{TR['avail'] + 1}" + ("".join(f"+{C}{BAL[q]['sweep'] + 1}" for q in prev))
            wds.write_formula(BAL[k]["sweep"], 3 + i, f"=-MIN({left},{C}{BAL[k]['open'] + 1}+{C}{BAL[k]['mand'] + 1})", f_n)
            prev.append(k)
        else:
            wds.write(BAL[k]["sweep"], 3 + i, 0, f_n)
    sw = "+".join(f"{C}{BAL[k]['sweep'] + 1}" for k in range(nT))
    wds.write_formula(TR["cash_close"], 3 + i, f"={C}{TR['cash_open'] + 1}+'Operating Model'!{C}{OR['Free cash flow after interest'] + 1}+{C}{TR['mand'] + 1}+{sw}", f_n)
    wds.write_formula(TR["debt"], 3 + i, "=" + "+".join(f"{C}{BAL[k]['close'] + 1}" for k in range(nT)), f_tot)
    wds.write_formula(TR["int"], 3 + i, "=" + "+".join(f"{C}{BAL[k]['intr'] + 1}" for k in range(nT)), f_n)
    wds.write_formula(TR["netdebt"], 3 + i, f"={C}{TR['debt'] + 1}-{C}{TR['cash_close'] + 1}", f_n)
    wds.write_formula(TR["lev"], 3 + i, f"={C}{TR['netdebt'] + 1}/'Operating Model'!{C}{OR['Adj. EBITDA'] + 1}", f_x)
    wds.write_formula(TR["cov"], 3 + i, f"='Operating Model'!{C}{OR['Adj. EBITDA'] + 1}/{C}{TR['int'] + 1}", f_x)
wb_int_row = TR["int"] + 1

# Returns
L = xl_col_to_name(2 + len(YEARS))
wr = sheet("Returns", "Returns and ability to pay (active scenario)"); wr.set_column("C:C", 14); wr.set_column("D:D", 58)
wr.write_formula("B4", '="Scenario: "&Inputs!$D$5', f_b)
RL = [("Exit adj. EBITDA (FY2031)", f"='Operating Model'!{L}{OR['Adj. EBITDA'] + 1}", f_lk, ""),
      ("Exit multiple", f"={R['exit_multiple']}", f_lkx, ""),
      ("Exit enterprise value", "=C6*C7", f_n, ""),
      ("(–) Net debt at exit", f"=-'Debt Schedule'!{L}{TR['netdebt'] + 1}", f_lk, ""),
      ("Exit equity value", "=C8+C9", f_tot, ""),
      ("Sponsor equity invested", f"={EQ_IN}", f_lk, "PIF rollover counted as invested at $210"),
      ("Holding period (years)", len(YEARS), fmt(font_color="#0000FF"), "Five full fiscal years FY2027-FY2031 (close Aug 2026)"),
      ("MoM", "=C10/C11", f_totx, ""),
      ("IRR", "=C13^(1/C12)-1", f_totp, ""),
      ("Equity for the target IRR", f"=C10/(1+{R['target_irr']})^C12", f_n, ""),
      ("Maximum price per share at the target IRR ($)",
       f"=(C15-{R['old_notes']}-{R['debt']}*{R['fin_fee_pct']}-{R['tlb']}*{R['oid_pct']}+{R['debt']}+{R['cash']}-{R['min_cash']})/({R['shares']}*(1+{R['advisory_pct']}))", f_totd, ""),
      ("vs offer price", f"=C16/{R['price']}-1", f_p, ""),
      ("vs unaffected price", f"=C16/{R['unaffected']}-1", f_p, ""),
      ("Exit multiple needed for the target IRR at the offer price", f"=(C11*(1+{R['target_irr']})^C12-C9)/C6", f_totx, "C9 is negative net debt"),
      ("Debt repaid over the hold", f"={R['debt']}-'Debt Schedule'!{L}{TR['debt'] + 1}", f_n, "")]
for k, (lab, form, f_, note) in enumerate(RL):
    wr.write(5 + k, 1, lab, f_b if f_ in (f_tot, f_totx, f_totp, f_totd) else None)
    (wr.write if isinstance(form, (int, float)) else wr.write_formula)(5 + k, 2, form, f_)
    if note: wr.write(5 + k, 3, note, f_note)
# Sensitivity
wsn = sheet("Sensitivity", "Sponsor IRR: entry price (rows) × exit multiple (columns)", "Active scenario · debt package fixed at $18bn, so the price changes only the equity cheque")
wsn.set_column("C:H", 11)
wsn.write(5, 1, "Price per share ↓ / exit multiple →", f_hdrl)
for j, m in enumerate(SENS_MULT): wsn.write(5, 2 + j, m, fmt(bold=True, font_color="#0000FF", bg_color="#ECEBE5", align="center", num_format='0.0"x"'))
for i, pr in enumerate(SENS_PRICE):
    wsn.write(6 + i, 1, pr, fmt(bold=True, font_color="#0000FF", num_format='$#,##0'))
    for j in range(len(SENS_MULT)):
        C = xl_col_to_name(2 + j)
        eq = (f"(($B{7 + i}*{R['shares']})*(1+{R['advisory_pct']})+{R['old_notes']}+{R['debt']}*{R['fin_fee_pct']}+{R['tlb']}*{R['oid_pct']}"
              f"-{R['debt']}-({R['cash']}-{R['min_cash']}))")
        wsn.write_formula(6 + i, 2 + j, f"=((Returns!$C$6*{C}$6+Returns!$C$9)/{eq})^(1/Returns!$C$12)-1", f_heat)
wsn.conditional_format(6, 2, 5 + len(SENS_PRICE), 1 + len(SENS_MULT), {"type": "cell", "criteria": ">=", "value": R["target_irr"],
                       "format": wb.add_format({"bg_color": "#1A1A18", "font_color": "#FFFFFF"})})
wsn.write(7 + len(SENS_PRICE), 1, "Dark cells: IRR at or above the 20% target", f_note)
wb.close()
assert TR['int'] + 1 == DS_INT_ROW
print("Saved", XLSX)

# %% [markdown]
# ## 10. Download

# %%
try:
    from google.colab import files
    for fpath in [XLSX, f"{OUTPUT_DIR}/deleveraging.png"]:
        files.download(fpath)
except ImportError:
    print("Outputs saved in:", os.path.abspath(OUTPUT_DIR))
