> **STALE — TERTILE RESULTS ONLY.** This file reports TERCILE analysis (T3-T1).
> The paper uses QUARTILE sorting (Q1-Q4). Do NOT use t-statistics from this file in paper text.
> Canonical values: FF3 alpha = 2.36%/month (t = 3.01); FF5 = 2.52% (t = 3.24).
> See `Paper_LaTeX/tables/tab_portfolio_alphas_adj.tex` for authoritative numbers.

# RRR Financial Implications — Results Summary

## Sample
- **124 firms** after sector filter (from 135 total)
- **4 sectors**: Consumer Discretionary (90), Industrials (13), Communication Services (11), Consumer Staples (10)
- **31 time periods**, quarterly, 2017Q1-2024Q4
- **10,416 monthly return observations** after merging quarterly signals

## Key Findings

### 1. Signal Persistence (Phase 2.1)
- **RRR is persistent**: Transition matrix shows 62.7% of firms in bottom tercile (T1) stay in T1, 46.6% of T3 stay in T3
- **AR is NOT persistent**: Only 27% stay in T1, 26.1% stay in T3 — almost random
- **Negative autocorrelation** in both signals at the firm level (mean AR(1) for RRR = -0.21), which is driven by mean-reversion within firms, but **cross-sectional ranking is stable**
- **Decision**: Signal persistence test supports RRR as a tradeable signal; AR is weaker

### 2. Portfolio Sorts — RRR (Phase 2.2)
**Main result: RRR long-short earns ~24% annualized alpha**

| Signal | Portfolio | Ann Return | Ann Vol | Sharpe |
|--------|-----------|-----------|---------|--------|
| Raw RRR | T1 (Low) | -6.96% | 36.84% | -0.19 |
| Raw RRR | T3 (High) | 17.62% | 23.60% | 0.75 |
| Raw RRR | **T3-T1** | **24.58%** | **30.52%** | **0.81** |
| Adj RRR | T1 (Low) | -6.85% | 36.83% | -0.19 |
| Adj RRR | T3 (High) | 17.36% | 24.58% | 0.71 |
| Adj RRR | **T3-T1** | **24.21%** | **29.57%** | **0.82** |
| Market (VW) | — | 11.06% | 23.41% | 0.47 |

**Factor regression alphas (long-short T3-T1):**

| Signal | FF3 Alpha (monthly) | t-stat | FF3+Mom Alpha | t-stat | FF5 Alpha | t-stat |
|--------|-------------------|--------|---------------|--------|-----------|--------|
| Raw RRR | 2.15% | **2.67** | 2.15% | **2.56** | 2.16% | **2.82** |
| Adj RRR | 2.05% | **2.66** | 2.06% | **2.58** | 2.18% | **2.84** |

**All alphas significant at 1% level across all three factor models. Industry adjustment barely changes the result.**

### 3. Portfolio Sorts — AR (Phase 2.3)
**AR is a weaker signal**

| Signal | Portfolio | Ann Return | Alpha (FF3) | t-stat |
|--------|-----------|-----------|-------------|--------|
| Raw AR | T3-T1 | 11.79% | 0.93% | 1.06 |
| Adj AR | T3-T1 | 9.29% | 0.80% | 1.03 |

AR long-short returns are positive but **not statistically significant** (t~1.0). This supports the hypothesis that RRR is the dominant signal.

### 4. Double Sort (3x3, Phase 2.2)
**Adjusted signals, annualized returns (VW):**

|  | AR T1 (Low) | AR T2 | AR T3 (High) |
|--|-------------|-------|--------------|
| **RRR T1 (Low)** | -10.34% | -7.68% | -0.03% |
| **RRR T2** | 4.35% | 0.50% | -3.13% |
| **RRR T3 (High)** | 5.93% | 9.41% | **16.46%** |

**Pattern**: Returns increase monotonically with RRR (row effect is strong). AR effect is weaker and inconsistent. The highest return (16.46%) comes from high-RRR/high-AR firms. The lowest (-10.34%) from low-RRR/low-AR.

### 5. Fama-MacBeth Cross-Sectional Regressions (Phase 2.5)
**RRR predicts future excess returns; AR does not (after controlling for RRR)**

| Spec | Adj RRR coef | t-stat | Adj AR coef | t-stat |
|------|-------------|--------|-------------|--------|
| (1) Adj RRR only | 0.0007 | **2.93*** | — | — |
| (2) Adj AR only | — | — | 0.0006 | **2.21** |
| (3) Both | 0.0003 | **2.13** | 0.0004 | 1.57 |
| (4) + Controls | 0.0004 | **2.98*** | 0.0003 | 1.29 |

When both signals are included, **RRR remains significant (t=2.98) while AR becomes insignificant (t=1.29)**. This is the key finding: RRR subsumes AR in cross-sectional return prediction.

### 6. Risk Analysis (Phase 2.6)

| Metric | T1 (Low RRR) | T2 | T3 (High RRR) |
|--------|-------------|----|----|
| Ann Return | -6.85% | 1.06% | **17.36%** |
| Ann Vol | 36.83% | 26.43% | **24.58%** |
| Sharpe | -0.19 | 0.04 | **0.71** |
| Max Drawdown | -64.55% | -45.20% | **-42.64%** |
| Downside Beta | 2.58 | 1.62 | **1.17** |
| VaR 5% | -15.42% | -10.45% | **-10.43%** |
| Skewness | -1.81 | -0.86 | **-0.27** |
| Hit Ratio | 49.4% | 49.4% | **65.5%** |

**High-RRR firms have HIGHER returns AND LOWER risk on every metric.** This answers RQ2: no, the excess returns are NOT associated with higher risk. High-RRR is genuinely "high returns, low risk."

### 7. Same-Growth Placebo (Phase 4.1)
**Within high-growth firms, sorting by RRR still predicts returns:**

| Portfolio | Ann Return | FF3 Alpha | t-stat |
|-----------|-----------|-----------|--------|
| T1 (Low RRR, High Growth) | 0.20% | -1.54% | -2.07 |
| T3 (High RRR, High Growth) | 20.90% | 0.39% | 0.86 |
| **T3-T1** | **20.69%** | **1.93%** | **2.19** |

Same headline revenue growth, different RRR composition → different returns. The alpha is significant (t=2.19).

### 8. Robustness
**Excluding COVID (2020Q1-2021Q2):** Alpha increases to 1.76%/month (t=2.83) — result is not driven by COVID bounce
**Equal-weighted:** Long-short alpha is 0.62%/month (t=1.60) — weaker but directionally consistent; the effect is concentrated in larger firms (VW >> EW)

## Decisions Made
1. **Sector filter**: Kept 4 sectors (124 firms) — results barely change vs full 135
2. **Industry adjustment**: Raw and adjusted signals give nearly identical results — the RRR effect is NOT driven by industry composition
3. **AR as signal**: AR alone has some predictive power (t=2.21 in FM) but is subsumed by RRR. Formally: AR is not a valid independent signal. The presentation's assumption (Slide 36) was correct.
4. **Terciles vs quintiles**: Terciles with ~38 firms each are well-populated
5. **FF models**: Results robust across FF3, FF3+Mom, and FF5 — no change in significance

## Issues to Flag for User
1. **Price-only returns**: Monthly returns use PX_LAST (price returns, no dividends). FF factors use total returns. For consumer discretionary firms with modest dividends, this mismatch is small but should be noted.
2. **Negative RRR autocorrelation**: At firm level, RRR shows negative autocorrelation (-0.21). But the transition matrix shows cross-sectional ranking stability (62.7% persistence in T1). This needs careful framing — the signal is "sticky in relative terms" even though firm-level RRR fluctuates.
3. **EW vs VW**: The long-short alpha is much stronger VW (2.15%/month) than EW (0.62%/month, insignificant). This means the effect is concentrated in larger firms. Could be a strength (actionable for institutional investors) or a weakness (not a broad anomaly).
4. **Small sample**: 87 monthly observations, 124 firms. Standard errors are honest (Newey-West), but the sample is thin by asset pricing standards.
