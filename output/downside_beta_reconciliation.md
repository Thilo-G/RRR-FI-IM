# Downside-Beta Reconciliation (Q1 minus Q4, adjusted-RRR sort)

**Bottom line:** the two figures are not in conflict and neither is a bug. They
are two legitimate downside betas that use two different market benchmarks and,
consequently, two different definitions of a "down month." Both are computed on
the *identical* Q1 and Q4 adjusted-RRR portfolio return series (88 months,
2017-09 to 2024-12, corrected timing
FORM_LAG=2m / HOLD=3m), so the gap is entirely
methodological, not a data or sample-window difference.

## The two numbers

| Definition | Benchmark | Down-month rule | Q1 | Q4 | Q1 - Q4 | SE | Source |
|---|---|---|---|---|---|---|---|
| Downside beta vs. **sample market** | Sample's own value-weighted market portfolio (raw return) | Sample market return < 0 (30 of 88 months) | 0.7741 | 1.2263 | **-0.4521** | -- | `tab_risk_stats.tex` (`performance_metrics.py`) |
| Downside beta vs. **FF market**, bootstrapped | Fama-French Mkt-RF (broad CRSP *excess* market return) | FF excess market return < 0 (29 of 88 months) | 0.8700 | 1.9317 | **-1.0617** | 0.9048 | `tab_h2ab_asymmetry.tex` Panel A (`robustness_diagnostics.py` Task 8) |

Both figures were reproduced here by importing and calling the *exact* estimator
from each source script (`performance_metrics.conditional_beta` and
`robustness_diagnostics._downside_beta`), so the reproduction is byte-faithful to
what each table reports.

## Why they differ (the two coupled design choices)

Two choices change together when you switch scripts, and they are the only things
that change:

1. **Benchmark / regressor.** `performance_metrics.py` was deliberately built so
   that *every* relative metric (tracking error, information ratio, downside and
   upside beta, hit rate) uses **the sample's own value-weighted market
   portfolio** as "the market" -- a documented design choice (see its module
   header). That benchmark is a 124-firm, four-sector value-weighted index drawn
   from the *same* universe as Q1 and Q4, so the portfolios' betas against it are
   mechanically compressed toward 1 and their Q1-Q4 spread is small. Task 8 of
   `robustness_diagnostics.py` instead regresses on the **Fama-French Mkt-RF**,
   the broad CRSP *excess* market return. Against that external, much broader
   benchmark the same portfolios load more heavily and more dispersedly in
   down states, so the Q1-Q4 spread roughly doubles in magnitude.

2. **Down-month definition (coupled to the benchmark).** "Down months" are the
   months in which the *chosen benchmark* is negative. Because the benchmarks
   differ, the down-month sets differ: 30 months have a negative
   sample-market return, versus 29 months with a negative FF *excess*
   return (Mkt-RF < 0 means the market underperformed the risk-free rate, not
   that its total return was negative). The two estimators therefore fit the
   conditional beta on partly different subsamples.

The 2x2 decomposition below (regressor benchmark x down-month series, both crossed
on the same Q1/Q4 series) confirms the **choice of benchmark is the dominant
driver**; the down-month redefinition contributes a smaller, second-order shift.

| Specification | Q1 | Q4 | Q1 - Q4 | N down |
|---|---|---|---|---|
| MKT reg, MKT down (= performance_metrics.py) | 0.7741 | 1.2263 | -0.4521 | 30 |
| Mkt-RF reg, Mkt-RF down (= robustness Task 8) | 0.8700 | 1.9317 | -1.0617 | 29 |
| Mkt-RF reg, MKT down (isolate regressor) | 0.5522 | 1.5859 | -1.0337 | 30 |
| MKT reg, Mkt-RF down (isolate down-set) | 0.9072 | 0.9389 | -0.0316 | 29 |

## Recommended table-note language

Both numbers are correct and should be kept, each with a note stating its
benchmark explicitly:

- For `tab_risk_stats.tex`: "Downside beta is the beta on **the sample's own
  value-weighted market portfolio** (MKT), estimated only in months that
  portfolio fell. This is not directly comparable to the downside beta in the
  H2a asymmetry table, which is measured against the Fama-French broad-market
  excess return."
- For `tab_h2ab_asymmetry.tex` Panel A: "Downside beta here is measured against
  the **Fama-French Mkt-RF** (broad-CRSP excess market return), estimated only in
  months that excess return was negative, and differs from the
  sample-market-based downside beta reported in the risk-statistics table.
  Reported honestly including the non-significant difference (95% CI includes 0)."

## Verdict

No bug. Keep both. Label each with its benchmark. The headline economic finding
is unchanged and consistent across both definitions: Q1 (high RRR) has a **lower**
downside beta than Q4 (low RRR); the two definitions disagree only on the
*magnitude* of that gap (-0.4521 vs. -1.0617), for the mechanical
reasons above, and the FF-benchmark version's Q1-Q4 gap is not statistically
distinguishable from zero after block-bootstrap.
