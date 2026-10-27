# Value-at-Risk-Analysis-Equibalanced-Portfolio-Amazon-AMZN-Apple-AAPL-
A comparison of five Value at Risk (VaR) methodologies on a $100 portfolio split equally between Amazon (AMZN) and Apple (AAPL) ($50 each). The analysis uses 6 months of daily closing prices (123 return observations) and is implemented in Excel VBA.


## Methodologies

1. **Parametric, flat sigma:** variance-covariance approach with volatility estimated over the full window.
2. **Parametric, EWMA:** RiskMetrics recursive volatility and covariance (λ = 0.94), more reactive to recent market conditions.
3. **Monte Carlo:** 10,000 correlated normal scenarios.
4. **Historical VaR:** empirical percentile of observed portfolio losses.
5. **Historical simulation:** past returns applied to the current portfolio value.

Joint portfolio risk: `σₚ = √[(W₁σ₁)² + (W₂σ₂)² + 2ρ·W₁σ₁·W₂σ₂]`

## Portfolio Statistics

| | AMZN | AAPL |
|---|---|---|
| Daily std dev | 1.96% | 1.33% |
| Correlation ρ | 0.2451 | |
| Portfolio daily σ | 1.31% | |

## Results (1-day joint VaR, $)

| Method | 95% | 99% | 99.5% |
|--------|------|------|------|
| Parametric (flat σ) | 2.16 | 3.06 | 3.38 |
| Parametric (EWMA) | 2.10 | 2.97 | 3.29 |
| Monte Carlo | 2.14 | 3.06 | 3.35 |
| Historical | 2.06 | 3.56 | 3.84 |

For a T-day horizon, parametric figures scale by √T.

## Key Findings

- **Sub-additivity holds in every method:** joint VaR is always below the sum of individual VaRs (at 99%, flat sigma: $3.06 vs $3.83, a diversification benefit of $0.77).
- **Parametric and Monte Carlo agree**, as both assume normal returns.
- **EWMA is slightly lower** than flat sigma, since recent volatility was below the 6-month average.
- **Historical VaR is higher at 99% and above**, which points to fat tails that the normal model underestimates.

## Files

- `report4_Basciu_Gabriele.xlsm`: VBA implementation and calculations
- Report (PDF): methodology, results and discussion

## Tech

Excel VBA, risk management, Monte Carlo simulation.
