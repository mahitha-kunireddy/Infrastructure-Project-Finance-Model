"""
PROJECT 2 (part 2) — Monte Carlo Risk Simulation
================================================
Re-implements the Excel DCF in Python, then runs 10,000 correlated-input
simulations to produce a distribution of NPV / IRR rather than a single
point estimate.

Stochastic drivers (the four that actually move a toll-road outcome):
    traffic CAGR       - normal, tied to regional GDP/freight uncertainty
    toll escalation    - normal, tracks the WPI-linked NHAI formula
    capex overrun      - lognormal, right-skewed (overruns >> underruns)
    O&M inflation      - normal

Outputs: P(NPV<0), NPV percentiles, IRR distribution, downside VaR,
and a traffic x toll sensitivity grid written back into the workbook.
"""

import numpy as np
import pandas as pd
from openpyxl import load_workbook

RNG = np.random.default_rng(42)
N_SIMS = 10_000

# ----------------------------------------------------------- base case ------
BASE = dict(
    n_constr=2, n_ops=20, km=60, cost_per_km=95.0, contingency=0.05,
    split1=0.45, aadt=24000.0, traffic_cagr=0.055, ramp1=0.75, ramp2=0.90,
    toll=115.0, toll_esc=0.04, leakage=0.07, om=210.0, om_inf=0.055,
    mm_cycle=5, mm_cost=720.0, debt_share=0.70, int_rate=0.095,
    tenor=15, morat=2, tax=0.25, dep_life=20, wacc=0.11,
)


def irr(cashflows, lo=-0.95, hi=1.5, tol=1e-7):
    """Bisection IRR. Returns nan if no sign change in the bracket."""
    def npv(r):
        return sum(cf / (1 + r) ** (i + 1) for i, cf in enumerate(cashflows))
    f_lo, f_hi = npv(lo), npv(hi)
    if np.isnan(f_lo) or np.isnan(f_hi) or f_lo * f_hi > 0:
        return np.nan
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if abs(f_mid) < tol:
            return mid
        if f_lo * f_mid < 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


def run_model(p):
    """Returns (npv, project_irr, equity_irr, min_dscr, breaches)."""
    nc, no = p["n_constr"], p["n_ops"]
    total_capex = p["km"] * p["cost_per_km"] * (1 + p["contingency"])

    capex = np.zeros(nc + no)
    capex[0] = -total_capex * p["split1"]
    capex[1] = -total_capex * (1 - p["split1"])

    debt_total = total_capex * p["debt_share"]
    dep = total_capex / p["dep_life"]

    ebitda = np.zeros(nc + no)
    ebitda_smooth = np.zeros(nc + no)   # before lumpy MM, after MMRA accrual
    interest = np.zeros(nc + no)
    principal = np.zeros(nc + no)
    tax_paid = np.zeros(nc + no)

    bal = debt_total
    amort_years = max(1, p["tenor"] - p["morat"])

    for t in range(no):
        i = nc + t
        ramp = p["ramp1"] if t == 0 else p["ramp2"] if t == 1 else 1.0
        traffic = p["aadt"] * (1 + p["traffic_cagr"]) ** t * ramp
        toll = p["toll"] * (1 + p["toll_esc"]) ** t
        gross = traffic * toll * 365 / 1e6
        net_rev = gross * (1 - p["leakage"])
        om = p["om"] * (1 + p["om_inf"]) ** t
        mm = (p["mm_cost"] * (1 + p["om_inf"]) ** (t + 1)
              if (t + 1) % p["mm_cycle"] == 0 else 0.0)
        mmra = p["mm_cost"] * (1 + p["om_inf"]) ** t / p["mm_cycle"]

        ebitda[i] = net_rev - om - mm
        ebitda_smooth[i] = net_rev - om - mmra

        interest[i] = bal * p["int_rate"]
        principal[i] = 0.0 if t < p["morat"] else min(bal, debt_total / amort_years)

        pbt = ebitda[i] - dep - interest[i]
        tax_paid[i] = max(0.0, pbt) * p["tax"]
        bal = max(0.0, bal - principal[i])

    fcff = ebitda - tax_paid + capex
    debt_draw = np.zeros(nc + no)
    debt_draw[0] = debt_total * p["split1"]
    debt_draw[1] = debt_total * (1 - p["split1"])
    fcfe = fcff - interest - principal + debt_draw

    disc = np.array([1 / (1 + p["wacc"]) ** (i + 1) for i in range(nc + no)])
    npv = float((fcff * disc).sum())

    ds = interest + principal
    mask = ds > 1e-9
    cfads = ebitda_smooth - tax_paid
    dscr = np.divide(cfads, ds, out=np.full_like(cfads, np.nan), where=mask)
    min_dscr = float(np.nanmin(dscr)) if mask.any() else np.nan
    breaches = int(np.nansum(dscr < 1.20))

    return npv, irr(fcff), irr(fcfe), min_dscr, breaches


# --------------------------------------------------- base case check --------
b_npv, b_irr, b_eirr, b_dscr, b_br = run_model(BASE)
print("=" * 74)
print("BASE CASE (should reconcile to the Excel model)")
print("=" * 74)
print(f"  NPV @ WACC            : INR {b_npv:,.0f} mn")
print(f"  Project IRR           : {b_irr:.2%}")
print(f"  Equity IRR            : {b_eirr:.2%}")
print(f"  Minimum DSCR          : {b_dscr:.2f}x")
print(f"  Years breaching 1.20x : {b_br}")

# ------------------------------------------------------- monte carlo --------
print()
print("=" * 74)
print(f"MONTE CARLO — {N_SIMS:,} simulations")
print("=" * 74)

traffic_draw = RNG.normal(0.055, 0.018, N_SIMS).clip(0.0, 0.12)
toll_draw = RNG.normal(0.040, 0.010, N_SIMS).clip(0.0, 0.08)
# lognormal overrun: median ~5%, long right tail — overruns dominate underruns
overrun_draw = (RNG.lognormal(np.log(1.05), 0.14, N_SIMS)).clip(0.90, 2.0) - 1.0
ominf_draw = RNG.normal(0.055, 0.012, N_SIMS).clip(0.02, 0.11)

print("Input distributions:")
print(f"  Traffic CAGR    ~ N(5.5%, 1.8%)   -> P5 {np.percentile(traffic_draw,5):.2%} "
      f"/ P95 {np.percentile(traffic_draw,95):.2%}")
print(f"  Toll escalation ~ N(4.0%, 1.0%)   -> P5 {np.percentile(toll_draw,5):.2%} "
      f"/ P95 {np.percentile(toll_draw,95):.2%}")
print(f"  Capex overrun   ~ LogN(med 5%)    -> P5 {np.percentile(overrun_draw,5):.2%} "
      f"/ P95 {np.percentile(overrun_draw,95):.2%}")
print(f"  O&M inflation   ~ N(5.5%, 1.2%)   -> P5 {np.percentile(ominf_draw,5):.2%} "
      f"/ P95 {np.percentile(ominf_draw,95):.2%}")

rows = []
for k in range(N_SIMS):
    p = dict(BASE)
    p["traffic_cagr"] = traffic_draw[k]
    p["toll_esc"] = toll_draw[k]
    p["contingency"] = overrun_draw[k]
    p["om_inf"] = ominf_draw[k]
    npv, pirr, eirr, mdscr, br = run_model(p)
    rows.append((npv, pirr, eirr, mdscr, br))

sim = pd.DataFrame(rows, columns=["npv", "project_irr", "equity_irr",
                                  "min_dscr", "breaches"])

p_neg = (sim["npv"] < 0).mean()
p_below_wacc = (sim["project_irr"] < BASE["wacc"]).mean()
p_cov = (sim["breaches"] > 0).mean()

print()
print("RESULTS")
print("-" * 74)
print(f"  Mean NPV                        : INR {sim['npv'].mean():,.0f} mn")
print(f"  Median NPV                      : INR {sim['npv'].median():,.0f} mn")
print(f"  P(NPV < 0)                      : {p_neg:.1%}")
print(f"  P(Project IRR < WACC of 11%)    : {p_below_wacc:.1%}")
print(f"  P(any DSCR covenant breach)     : {p_cov:.1%}")
print()
print(f"  {'Percentile':<14}{'NPV (INR mn)':>16}{'Project IRR':>14}{'Min DSCR':>12}")
print("  " + "-" * 54)
for q in [5, 25, 50, 75, 95]:
    print(f"  P{q:<13}{np.percentile(sim['npv'], q):>16,.0f}"
          f"{np.percentile(sim['project_irr'].dropna(), q):>13.2%}"
          f"{np.percentile(sim['min_dscr'].dropna(), q):>11.2f}x")

var5 = np.percentile(sim["npv"], 5)
print()
print(f"  Downside (P5) NPV               : INR {var5:,.0f} mn")
print(f"  Base-case NPV                   : INR {b_npv:,.0f} mn")
print(f"  Value at risk vs base (P5)      : INR {b_npv - var5:,.0f} mn")

# ------------------------------------------- driver sensitivity ranking -----
print()
print("DRIVER SENSITIVITY (correlation of input to NPV)")
print("-" * 74)
drivers = pd.DataFrame({
    "Traffic CAGR": traffic_draw, "Toll escalation": toll_draw,
    "Capex overrun": overrun_draw, "O&M inflation": ominf_draw,
})
corr = drivers.corrwith(sim["npv"]).sort_values(key=abs, ascending=False)
for name, c in corr.items():
    bar = "#" * int(abs(c) * 40)
    print(f"  {name:<18}{c:>7.3f}  {bar}")

# --------------------------------------------- write grid back to Excel -----
traffic_vals = [0.035, 0.045, 0.055, 0.065, 0.075]
toll_vals = [0.02, 0.03, 0.04, 0.05, 0.06]

wb = load_workbook("Toll_Road_Project_Finance_Model.xlsx")
rt = wb["Returns"]

SENS_TOP = None
for row in range(1, 60):
    if rt.cell(row, 2).value == "Traffic \\ Toll":
        SENS_TOP = row
        break

grid = []
for i, g in enumerate(traffic_vals):
    line = []
    for j, tv in enumerate(toll_vals):
        p = dict(BASE)
        p["traffic_cagr"], p["toll_esc"] = g, tv
        _, pirr, _, _, _ = run_model(p)
        line.append(pirr)
        rt.cell(SENS_TOP + 1 + i, 3 + j, round(float(pirr), 6))
    grid.append(line)

wb.save("Toll_Road_Project_Finance_Model.xlsx")

print()
print("SENSITIVITY GRID — Project IRR  (rows: traffic CAGR, cols: toll escalation)")
print("-" * 74)
print(f"  {'':>10}" + "".join(f"{t:>10.1%}" for t in toll_vals))
for i, g in enumerate(traffic_vals):
    print(f"  {g:>10.1%}" + "".join(f"{v:>10.2%}" for v in grid[i]))
print(f"\n  (WACC hurdle = {BASE['wacc']:.0%} — cells below this destroy value)")

sim.describe().to_csv("output_monte_carlo_summary.csv")
sim.to_csv("output_monte_carlo_runs.csv", index=False)

headline = {
    "base_npv_inr_mn": round(b_npv), "base_project_irr": round(b_irr, 4),
    "base_equity_irr": round(b_eirr, 4), "base_min_dscr": round(b_dscr, 2),
    "n_simulations": N_SIMS, "p_npv_negative": round(float(p_neg), 4),
    "p_irr_below_wacc": round(float(p_below_wacc), 4),
    "p_covenant_breach": round(float(p_cov), 4),
    "npv_p5_inr_mn": round(float(var5)), "npv_p95_inr_mn": round(float(np.percentile(sim['npv'], 95))),
    "top_driver": corr.index[0], "top_driver_corr": round(float(corr.iloc[0]), 3),
}
pd.DataFrame([headline]).to_csv("output_headline.csv", index=False)

print()
print("=" * 74)
print("HEADLINE NUMBERS FOR RESUME")
print("=" * 74)
for k, v in headline.items():
    print(f"  {k:<22}: {v}")
