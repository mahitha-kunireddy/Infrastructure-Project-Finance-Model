# Infrastructure Project Finance Model — Toll Road BOT Concession

A 22-year project finance model for a toll road concession, built as a fully formula-driven Excel workbook with a Python Monte Carlo risk layer.

**Tech:** Excel (openpyxl, 540 live formulas) · Python (NumPy, pandas) · DCF & project finance

---

## What this does

Models a 60km four-lane toll road built under a BOT (Build-Operate-Transfer) concession — two years of construction followed by twenty years of tolled operations — and answers two questions:

1. **Is it worth building?** NPV, project IRR, equity IRR, payback[cite: 5]
2. **Is it financeable?** DSCR against lender covenants, and how likely the returns are to disappoint[cite: 5]

The Excel workbook is entirely formula-driven — every output traces back to a labelled input cell, and the model recalculates end to end when any assumption changes[cite: 5]. The Python layer re-implements the same logic to run 10,000 simulations across uncertain inputs[cite: 5].

### Project finance mechanics included

Features that materially change the answer and are usually missing from simplified DCF models[cite: 5]:

- **Traffic ramp-up curve** — 75% / 90% / full utilisation over the first three operating years[cite: 5]
- **Major Maintenance Reserve Account** — annual accrual smoothing lumpy resurfacing costs for covenant testing[cite: 5]
- **Principal repayment moratorium** — a debt holiday during ramp-up, standard in Indian BOT financing[cite: 5]
- **DSCR on smoothed CFADS** — tested against a lender covenant floor, with breach counting[cite: 5]

---

## Results

### Base case

| Metric | Value |
|---|---|
| Total project cost | ₹5,985mn |
| **NPV @ 11% WACC** | **+₹2,310mn** |
| **Project IRR (FCFF)** | **15.26%** |
| **Equity IRR (FCFE)** | **20.59%** |
| Spread over WACC | +4.26% |
| Discounted payback | Year 16 |

Equity IRR exceeds project IRR because the 9.5% cost of debt sits below the unlevered project return — positive financial leverage[cite: 5].

### The finding that matters

| Metric | Value |
|---|---|
| Minimum DSCR | **0.88x** |
| Average DSCR | 2.51x |
| Lender covenant floor | 1.20x |
| **Operating years breaching covenant** | **3** |

The project is NPV-positive but **not bankable as structured**[cite: 5]. Debt service coverage falls below a standard 1.20x covenant in three operating years during traffic ramp-up[cite: 5]. The structure would need a debt service reserve account or an extended moratorium before a lender would take it[cite: 5].

### Monte Carlo — 10,000 simulations

Stochastic inputs: traffic CAGR ~N(5.5%, 1.8%) · toll escalation ~N(4.0%, 1.0%) · capex overrun ~LogNormal (right-skewed, since overruns dominate underruns) · O&M inflation ~N(5.5%, 1.2%)[cite: 5]

| Outcome | Probability |
|---|---|
| **NPV < 0** | **8.3%** |
| Project IRR < WACC | 4.8% |

| Percentile | NPV (₹mn) | Project IRR |
|---|---|---|
| P5 (downside) | -400 | 11.00% |
| P50 (median) | 2,219 | 15.27% |
| P95 (upside) | 5,740 | 19.63% |

Value at risk against base case at P5: **₹2,711mn**[cite: 5].

### Driver sensitivity

Correlation of each input to NPV across the simulation[cite: 5]:

| Driver | Correlation |
|---|---|
| **Traffic CAGR** | **+0.804** |
| Toll escalation | +0.448 |
| Capex overrun | -0.321 |
| O&M inflation | -0.148 |

Traffic growth dominates everything else combined[cite: 5]. Capex overrun — the risk that gets the most attention in practice — matters materially less than traffic[cite: 5].

### Sensitivity grid — Project IRR

Rows: traffic CAGR · Columns: toll escalation · WACC hurdle 11%[cite: 5]

| | 2.0% | 3.0% | 4.0% | 5.0% | 6.0% |
|---|---|---|---|---|---|
| **3.5%** | 10.11% | 11.53% | 12.83% | 14.07% | 15.25% |
| **4.5%** | 11.51% | 12.82% | 14.07% | 15.26% | 16.41% |
| **5.5%** | 12.79% | 14.05% | **15.26%** | 16.42% | 17.55% |
| **6.5%** | 14.01% | 15.23% | 16.40% | 17.55% | 18.67% |
| **7.5%** | 15.18% | 16.37% | 17.52% | 18.65% | 19.76% |

The project only fails its hurdle rate in the bottom-left corner — weak traffic *and* weak tariff escalation together[cite: 5].

---

## Assumptions

**All operating assumptions in this model are illustrative placeholders**, chosen to sit within plausible ranges for an Indian highway concession[cite: 5]. They are flagged as such inside the workbook[cite: 5]. This repository demonstrates the model structure and analytical method — to apply it to a specific asset, replace the inputs with figures sourced from that project's concession agreement, NHAI bid documents, or the developer's disclosures[cite: 5].

Editable input cells are marked in blue; key return drivers are highlighted in yellow[cite: 5].

---

## Running it

```bash
pip install openpyxl numpy pandas
python3 build_dcf.py          # generates the Excel workbook
python3 run_monte_carlo.py    # runs 10,000 simulations, fills the sensitivity grid
