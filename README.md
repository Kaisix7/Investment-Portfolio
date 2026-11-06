# Investment Portfolio Optimization & Efficient Frontier Analysis

[![Excel](https://img.shields.io/badge/Model-Microsoft_Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)](https://github.com/Kaisix7/Investment-Portfolio/blob/main/Investment%20Portfolio.xlsx)
[![Methodology](https://img.shields.io/badge/Theory-Modern_Portfolio_Theory_MPT-0284c7?style=flat)](https://en.wikipedia.org/wiki/Modern_portfolio_theory)
[![Optimization](https://img.shields.io/badge/Optimization-Excel_Solver-f59e0b?style=flat)](https://www.solver.com/)
[![License](https://img.shields.io/badge/License-MIT-emerald?style=flat)](LICENSE)

An institutional-grade portfolio analysis and risk-return optimization model constructed in Microsoft Excel, implementing **Harry Markowitz's Modern Portfolio Theory (MPT)**. The model analyzes 5 years of daily market data (1,254 trading days from October 2021 to October 2026) across five prominent U.S. equities against the S&P 500 benchmark to derive the **Global Minimum Variance (GMV) Portfolio** and trace the **Efficient Frontier**.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Objectives](#objectives)
- [Assets Analyzed](#assets-analyzed)
- [Financial Methodology](#financial-methodology)
- [Portfolio Allocation](#portfolio-allocation)
- [Risk and Return Analysis](#risk-and-return-analysis)
- [Efficient Frontier](#efficient-frontier)
- [Tools and Technologies](#tools-and-technologies)
- [Key Findings](#key-findings)
- [Repository Structure](#repository-structure)
- [Workbook Guide](#workbook-guide)
- [License](#license)

---

## Project Overview

In modern quantitative finance, constructing an investment portfolio requires more than simply selecting high-performing stocks; it demands optimizing the trade-off between risk and expected return. By exploiting the mathematical covariance between disparate asset classes, investors can engineer a diversified portfolio that delivers a target expected return while minimizing total volatility.

This repository contains the complete quantitative workbook [`Investment Portfolio.xlsx`](Investment%20Portfolio.xlsx), featuring daily adjusted price observations, daily discrete return computations, variance-covariance matrices, linear regression betas against the S&P 500, utility maximization under risk aversion, and non-linear quadratic programming executed via the **Excel Solver** add-in.

![Efficient Frontier](images/efficient-frontier.png)

---

## Objectives

1. **Empirical Return & Risk Profiling**: Calculate historical expected daily returns, variance, standard deviation, and CAPM Beta for each individual equity over a 5-year investment horizon (1,254 trading days).
2. **Covariance & Correlation Modeling**: Formulate the full $5 \times 5$ population variance-covariance matrix ($\Sigma$) to quantify pairwise co-movements and identify diversification drivers.
3. **Quadratic Portfolio Optimization**: Use non-linear mathematical optimization in Excel Solver to solve for the **Global Minimum Variance (GMV) Portfolio** under full investment ($\sum w_i = 1$) and long-only ($w_i \ge 0$) constraints.
4. **Efficient Frontier & Benchmark Comparison**: Evaluate the resulting optimal portfolio against the S&P 500 benchmark and examine the Capital Allocation Line (CAL) under investor risk-aversion parameters.
5. **Practical Capital Allocation Assessment**: Inspect utility-based capital allocation between risky and risk-free assets (U.S. Treasury Bills).

---

## Assets Analyzed

The investment universe consists of five major U.S. large-cap corporations spanning growth technology, consumer staples, and retail sectors, alongside the S&P 500 (`^GSPC`) as the market benchmark:

| Ticker | Company Name | Sector | Business Model Focus |
|:---|:---|:---|:---|
| **AAPL** | **Apple Inc.** | Information Technology | Consumer electronics, software ecosystems, subscription services |
| **TSLA** | **Tesla, Inc.** | Consumer Discretionary / Tech | Electric vehicles, energy storage, autonomous driving |
| **KO** | **The Coca-Cola Company** | Consumer Staples | Global non-alcoholic beverage manufacturing and distribution |
| **MCD** | **McDonald's Corporation** | Consumer Discretionary | Global quick-service restaurant franchising and real estate |
| **WMT** | **Walmart Inc.** | Consumer Staples | Omnichannel hypermarket retail and supply-chain logistics |
| **^GSPC** | **S&P 500 Index** | Market Benchmark | U.S. broad equity market benchmark (500 large-cap equities) |

---

## Financial Methodology

The mathematical foundation follows classical **Markowitz Modern Portfolio Theory (1952)** and the **Capital Asset Pricing Model (CAPM)**:

### 1. Daily Arithmetic Returns
Discrete daily percentage returns are calculated from adjusted closing prices:

$$R_{i,t} = \frac{P_{i,t}}{P_{i,t-1}} - 1$$

In Excel:
```excel
=(B3/B4)-1
```

### 2. Expected Return & Volatility
- **Expected Daily Return**:
  $$E(R_i) = \frac{1}{T} \sum_{t=1}^{T} R_{i,t} \quad \longrightarrow \quad \text{Excel: } =AVERAGE(H3:H1256)$$
- **Sample Variance & Standard Deviation**:
  $$\sigma_i^2 = \frac{1}{T-1} \sum_{t=1}^{T} (R_{i,t} - E(R_i))^2 \quad \longrightarrow \quad \text{Excel: } =VAR.S(H3:H1256)$$
  $$\sigma_i = \sqrt{\sigma_i^2} \quad \longrightarrow \quad \text{Excel: } =STDEV.S(H3:H1256)$$

### 3. Systematic Risk (Beta $\beta$)
Beta measures an asset's systematic risk sensitivity relative to the market benchmark (S&P 500):

$$\beta_i = \frac{\text{Cov}(R_i, R_M)}{\sigma_M^2} = \text{SLOPE}(R_i, R_M)$$

In Excel:
```excel
=SLOPE(H3:H1256, $M$3:$M$1256)
```

### 4. Variance-Covariance Matrix ($\mathbf{\Sigma}$)
The $5 \times 5$ symmetric variance-covariance matrix is formulated using daily returns ($T = 1,253$ return steps):
- **Diagonal Elements ($\sigma_i^2$, for $i = j$)**: Represent the **population variance** of asset $i$, calculated in Excel using:
  $$\sigma_i^2 = \text{Var}_P(R_i) = \frac{1}{T} \sum_{t=1}^T (R_{i,t} - \mu_i)^2 \quad \longrightarrow \quad \text{Excel: } =VARP(ПОРТФЕЛЬ!\$H\$3:\$H\$1255)$$
  *(Note: In Excel, `=VARP` computes single-asset population variance, not covariance between two distinct series).*
- **Off-Diagonal Elements ($\sigma_{ij}$, for $i \ne j$)**: Represent the pairwise **population covariance** between assets $i$ and $j$:
  $$\sigma_{ij} = \text{Cov}_P(R_i, R_j) = \frac{1}{T} \sum_{t=1}^T (R_{i,t} - \mu_i)(R_{j,t} - \mu_j) \quad \longrightarrow \quad \text{Excel: } =COVARIANCE.P(\dots)$$
  In the workbook, the lower-triangular cells contain the evaluated population covariances, while the upper-triangular cells mirror them via direct cell references (e.g., cell `S13` is `=R14`) to ensure exact matrix symmetry.

### 5. Annualized Returns Terminology
Quantitative portfolio modeling distinguishes between arithmetic expected return and geometric compounded return:
- **Annualized Arithmetic Expected Return**:
  $$E(R)_{\text{arith, ann}} = E(R_{\text{daily}}) \times 252$$
  This linear scaling represents the annual expected arithmetic return assumed in single-period Markowitz Modern Portfolio Theory (MPT), since mathematical expectation is a linear operator: $E\left[\sum_{t=1}^{252} R_t\right] = 252 \times E[R_t]$.
- **Compounded Annual Growth Rate (CAGR) / Geometric Return**:
  $$\text{CAGR} = \left(\frac{P_T}{P_0}\right)^{\frac{252}{T}} - 1 \approx E(R)_{\text{arith, ann}} - \frac{1}{2}\sigma_{\text{ann}}^2$$
  This reflects true multi-year capital compounding. The difference between arithmetic return and CAGR is driven by **volatility drag** ($\approx \frac{1}{2}\sigma^2$), which heavily penalizes high-volatility assets.

### 6. Portfolio Variance & Matrix Formulation
Using matrix multiplication, portfolio variance for weight vector $\mathbf{w} = [w_1, \dots, w_5]^T$ is:

$$\sigma_p^2 = \mathbf{w}^T \mathbf{\Sigma} \mathbf{w} = \sum_{i=1}^n \sum_{j=1}^n w_i w_j \sigma_{ij}$$

In Excel:
```excel
=MMULT(TRANSPOSE(W13:W17), MMULT(R13:V17, W13:W17))
```

### 7. Portfolio Expected Return & Beta
$$\mu_p = \sum_{i=1}^n w_i E(R_i) \quad \longrightarrow \quad \text{Excel: } =SUMPRODUCT(W13:W17, X13:X17)$$
$$\beta_p = \sum_{i=1}^n w_i \beta_i \quad \longrightarrow \quad \text{Excel: } =W13*R7 + W14*S7 + W15*T7 + W16*U7 + W17*V7$$

### 8. Global Minimum Variance Optimization Formulation
Excel Solver was executed with the following objective and constraints:
- **Objective**: Minimize $\sigma_p^2$ (Cell `R19`)
- **By Changing Variable Cells**: Weights $\mathbf{w}$ (Cells `W13:W17`)
- **Constraints**:
  1. Full investment: $\sum_{i=1}^5 w_i = 1$ (Cell `W18 = 1`)
  2. Long-only (no short sales): $w_i \ge 0 \quad \forall i$ (Cells `W13:W17 >= 0`)

---

## Portfolio Allocation

The optimization routine converged on the following optimal asset allocation:

![Portfolio Allocation](images/portfolio-allocation.png)

| Asset | Ticker | Optimal Weight ($w_i$) | Sector | Rationale in Optimization |
|:---|:---:|:---:|:---|:---|
| **Coca-Cola** | **KO** | **41.58%** | Consumer Staples | Lowest standalone volatility ($\sigma = 1.07\%$) and defensive cash flows |
| **McDonald's** | **MCD** | **30.50%** | Consumer Discretionary | Low beta ($\beta = 0.36$) and stable consumer demand |
| **Walmart** | **WMT** | **17.79%** | Consumer Staples | Strong risk-adjusted returns with moderate variance |
| **Apple** | **AAPL** | **7.94%** | Technology | Core mega-cap growth providing return enhancement |
| **Tesla** | **TSLA** | **2.19%** | Consumer / Tech | Minimized due to extreme standalone volatility ($\sigma = 3.76\%$) |
| **Total** | | **100.00%** | | Fully invested long-only portfolio |

### Why Consumer Staples Dominate the Allocation
In a minimum-variance objective, the optimizer allocates heavily towards assets with **low individual variance** and **low correlation** to the rest of the basket:
- **Coca-Cola (41.58%)** and **McDonald's (30.50%)** combine for **72.08%** of the portfolio because their defensive revenues insulate them from macroeconomic shocks.
- **Tesla (2.19%)** has the highest volatility ($\sigma = 3.76\%$ daily, $\approx 59.7\%$ annualized) and high beta ($\beta = 2.07$), leading Solver to severely constrain its allocation to preserve low portfolio variance.

---

## Risk and Return Analysis

### Summary Statistics of Analyzed Assets (5-Year Horizon)

| Asset / Benchmark | Daily Return $E(R)$ | Annualized Arithmetic Return ($E(R) \times 252$) | 5-Year CAGR (Compounded) | Daily Variance $\sigma^2$ | Daily Std Dev $\sigma$ | Annualized Volatility ($\sigma \times \sqrt{252}$) | CAPM Beta $\beta$ | Weight ($w$) |
|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **Apple (AAPL)** | 0.0833% | 20.99% | 18.53% | 0.00031163 | 1.7653% | 28.02% | 1.1741 | 7.94% |
| **Tesla (TSLA)** | 0.0773% | 19.48% | 1.87% | 0.00141450 | 3.7610% | 59.70% | 2.0749 | 2.19% |
| **Coca-Cola (KO)** | 0.0485% | 12.21% | 11.29% | 0.00011364 | 1.0660% | 16.92% | 0.2473 | 41.58% |
| **McDonald's (MCD)** | 0.0101% | 2.54% | 1.00% | 0.00012882 | 1.1350% | 18.02% | 0.3575 | 30.50% |
| **Walmart (WMT)** | 0.0799% | 20.15% | 19.43% | 0.00020308 | 1.4251% | 22.62% | 0.4035 | 17.79% |
| **S&P 500 (^GSPC)** | 0.0466% | 11.74% | 10.82% | 0.00011451 | 1.0701% | 16.99% | 1.0000 | *Benchmark* |
| **GMV Portfolio** | **0.0458%** | **11.53%** | **10.52%** | **0.00008013** | **0.8952%** | **14.21%** | **0.4223** | **100.00%** |

*Methodological Clarification*:
- **Annualized Arithmetic Return** ($E(R) \times 252$) is the linear expected annual return modeled in classical Modern Portfolio Theory.
- **Compounded Annual Growth Rate (CAGR)** ($(P_T / P_0)^{252/T} - 1$) reflects actual realized geometric multi-year compounding. Notice the massive gap for Tesla ($19.48\%$ arithmetic vs. $1.87\%$ CAGR), caused by extreme volatility drag ($\approx \frac{1}{2}\sigma_{\text{ann}}^2 \approx 17.8\%$).
- **Annualized Volatility** scales daily standard deviation by the square root of time: $\sigma_{\text{ann}} = \sigma_{\text{daily}} \times \sqrt{252}$.

### Correlation Matrix ($\rho_{ij}$)

The pairwise correlation matrix demonstrates why diversification is so potent across these assets:

| Asset | Apple (AAPL) | Tesla (TSLA) | Coca-Cola (KO) | McDonald's (MCD) | Walmart (WMT) |
|:---|:---:|:---:|:---:|:---:|:---:|
| **Apple** | **1.0000** | 0.4625 | 0.2455 | 0.2936 | 0.2571 |
| **Tesla** | 0.4625 | **1.0000** | **0.0016** | **0.0880** | **0.1365** |
| **Coca-Cola** | 0.2455 | **0.0016** | **1.0000** | 0.5310 | 0.3557 |
| **McDonald's** | 0.2936 | **0.0880** | 0.5310 | **1.0000** | 0.3005 |
| **Walmart** | 0.2571 | **0.1365** | 0.3557 | 0.3005 | **1.0000** |

> **Key Observation**: The correlation between Tesla and Coca-Cola is virtually zero ($\rho = 0.0016$), and Tesla's correlations with McDonald's ($\rho = 0.0880$) and Walmart ($\rho = 0.1365$) are extremely low. This near-orthogonality eliminates unsystematic variance when combined in a portfolio.

---

## Efficient Frontier

### Mathematical Formulation & Unconstrained vs. Long-Only Frontier

For a given vector of expected returns $\mathbf{\mu}$ and variance-covariance matrix $\mathbf{\Sigma}$, the analytical minimum-variance portfolio for any target expected return $r$ under unconstrained weights satisfies the Markowitz hyperbola:

$$\sigma^2(r) = \frac{A r^2 - 2B r + C}{\Delta}$$

where:
- $A = \mathbf{1}^T \mathbf{\Sigma}^{-1} \mathbf{1} \approx 12,479.76$
- $B = \mathbf{1}^T \mathbf{\Sigma}^{-1} \mathbf{\mu} \approx 5.710142$
- $C = \mathbf{\mu}^T \mathbf{\Sigma}^{-1} \mathbf{\mu} \approx 0.005674$
- $\Delta = AC - B^2 \approx 38.2000$

#### Global Minimum Variance (GMV) Portfolio
The vertex of this hyperbola represents the Global Minimum Variance portfolio:
- **GMV Expected Return**:
  $$\mu_{GMV} = \frac{B}{A} = \frac{5.710142}{12,479.76} \approx 0.04576\% \text{ daily} \quad (\approx 11.53\% \text{ annualized})$$
- **GMV Portfolio Variance**:
  $$\sigma_{GMV}^2 = \frac{1}{A} = \frac{1}{12,479.76} \approx 8.01298 \times 10^{-5}$$
- **GMV Portfolio Standard Deviation**:
  $$\sigma_{GMV} = \frac{1}{\sqrt{A}} = \frac{1}{\sqrt{12,479.76}} \approx 0.0089515 = 0.89515\% \text{ daily} \quad (\approx 14.21\% \text{ annualized})$$
  *(Note: Portfolio variance is $\sigma^2 = 1/A$, while portfolio standard deviation is $\sigma = 1/\sqrt{A}$).*

#### Analytical Frontier vs. Long-Only Constraints ($w_i \ge 0$)
A crucial mathematical finding from this audit concerns the compatibility between the analytical formula and the long-only constraints ($w_i \ge 0$):
1. **At the GMV Vertex**: The analytical weight vector is:
   $$\mathbf{w}_{GMV} = \frac{\mathbf{\Sigma}^{-1}\mathbf{1}}{A} = [7.94\%, 2.19\%, 41.58\%, 30.50\%, 17.79\%]^T$$
   Because every constituent weight is strictly positive, the unconstrained GMV portfolio **naturally satisfies the long-only condition without any active constraints**. Thus, the analytical solution and the Excel Solver constrained solution are mathematically identical to 8 decimal places.
2. **Along the Remainder of the Frontier**: As the target return $r$ varies along the unconstrained analytical frontier $\mathbf{w}(r) = \mathbf{g} + \mathbf{h}r$:
   - For high target returns ($r > 0.065\%$ daily), the optimal weight for **McDonald's becomes negative** (e.g., $w_{MCD} = -24.79\%$ at $r = 0.080\%$).
   - For low target returns ($r < 0.035\%$ daily), **Apple and Walmart weights become negative** (e.g., $w_{AAPL} = -10.84\%$ at $r = 0.010\%$).
   - Consequently, the continuous analytical hyperbola **assumes short sales are permitted** and is **not compatible with long-only constraints** across the entire curve. Tracing a true long-only frontier requires solving a quadratic programming problem with inequality constraints ($w_i \ge 0$) for each target return level.

### Excel Solver Implementation Scope
- **Verified Implementation**: In [`Investment Portfolio.xlsx`](Investment%20Portfolio.xlsx), **Excel Solver was configured and executed specifically for the Global Minimum Variance (GMV) Portfolio** (minimizing cell `R19` subject to $\sum w_i = 1$ and $w_i \ge 0$).
- **Multi-Point Table Architecture**: The worksheet includes column headers and reserved rows (`Q26:R36`) for an Efficient Frontier table and an associated chart object. However, individual intermediate frontier points were not solved via Solver in the workbook. The chart in `images/efficient-frontier.png` plots the continuous Markowitz frontier alongside the verified GMV Solver optimum.

---

## Tools and Technologies

- **Microsoft Excel**: Core numerical modeling, dynamic array computation, and formula architecture.
- **Excel Solver Add-in**: Non-linear GRG (Generalized Reduced Gradient) engine utilized for quadratic variance minimization under linear equality and inequality constraints.
- **Matrix Functions**:
  - `=MMULT(TRANSPOSE(w), MMULT(Cov, w))` for portfolio variance calculation.
  - `=SUMPRODUCT(w, r)` for portfolio expected return.
  - `=SLOPE(...)` for CAPM linear regression beta coefficients.
  - `=VAR.S(...)`, `=STDEV.S(...)`, and `=VARP(...)` for statistical sample and population dispersion.
- **Python**: Verification of analytical matrix inversion, Kuhn-Tucker optimality conditions, and visual chart rendering.

---

## Key Findings

1. **The Diversification Dividend**:
   The optimized portfolio achieves a daily standard deviation of **$0.8952\%$** ($14.21\%$ annualized), which is **lower than every single constituent asset** (Coca-Cola at $1.066\%$, McDonald's at $1.135\%$, Walmart at $1.425\%$, Apple at $1.765\%$, and Tesla at $3.761\%$). This provides concrete mathematical evidence of Markowitz's diversification dividend.

2. **Market-Beating Volatility with Benchmark-Matching Return**:
   - The S&P 500 benchmark yielded **$11.74\%$** annualized return with **$16.99\%$** annualized volatility.
   - The GMV Portfolio yielded **$11.53\%$** annualized return with only **$14.21\%$** annualized volatility.
   - The portfolio captures **$\sim 98\%$ of the S&P 500's return** with **$\sim 16\%$ lower volatility** and a **Beta of only $0.4223$**.

3. **Asset Volatility Asymmetry**:
   Tesla exhibited more than $3.5\times$ the daily standard deviation of Coca-Cola ($3.761\%$ vs $1.066\%$). In a variance-minimizing mandate, even high expected returns cannot overcome quadratic volatility penalties, resulting in Tesla receiving a minimal allocation ($2.19\%$).

4. **Sharpe Ratio Frequency Consistency & Audit**:
   - In Excel cell `R22`, the formula is written as `=(R21-R23)/R20`, which subtracts the annualized 2.37% Treasury yield (Cell `R23`) directly from the daily expected return ($0.0458\%$, Cell `R21`), dividing by daily standard deviation ($0.8952\%$, Cell `R20`). This produces a nominal $-2.5965$ due to dimensional frequency mismatch.
   - **Mathematically Consistent Daily Sharpe Ratio**:
     $$R_{f,\text{daily}} = \frac{2.37\%}{252} \approx 0.009405\%$$
     $$\text{Sharpe}_{\text{daily}} = \frac{E(R_{p,\text{daily}}) - R_{f,\text{daily}}}{\sigma_{p,\text{daily}}} = \frac{0.00045755 - 0.00009405}{0.00895152} \approx \mathbf{0.0406}$$
   - **Mathematically Consistent Annualized Sharpe Ratio**:
     $$\text{Sharpe}_{\text{ann}} = \frac{E(R_{p,\text{ann}}) - R_{f,\text{ann}}}{\sigma_{p,\text{ann}}} = \frac{11.5303\% - 2.3700\%}{14.2101\%} \approx \mathbf{0.6446}$$
     Equivalently, via square-root-of-time scaling:
     $$\text{Sharpe}_{\text{ann}} = \text{Sharpe}_{\text{daily}} \times \sqrt{252} \approx 0.040608 \times \sqrt{252} \approx \mathbf{0.6446}$$
   - *Workbook Preservation Note*: In accordance with preserving the original submission workbook, the Excel formulas have been maintained as originally calculated, and the frequency-consistent metrics are fully detailed here for quantitative audit purposes.

---

## Repository Structure

```text
Investment Portfolio/
│
├── README.md                  # Comprehensive professional project documentation
├── Investment Portfolio.xlsx   # Original Excel workbook with all formulas & Solver models
├── images/
│   ├── efficient-frontier.png # Publication-grade Efficient Frontier & CAL chart
│   └── portfolio-allocation.png # Optimal asset allocation donut breakdown chart
└── LICENSE                    # Open-source MIT License
```

---

## Workbook Guide

To explore or extend the calculations in Microsoft Excel:

1. Download or clone this repository:
   ```bash
   git clone https://github.com/Kaisix7/Investment-Portfolio.git
   ```
2. Open [`Investment Portfolio.xlsx`](Investment%20Portfolio.xlsx) directly in Microsoft Excel (or view it via GitHub's integrated spreadsheet viewer).
3. Navigate to worksheet **`ПОРТФЕЛЬ`**:
   - **Columns A–F**: Raw historical daily price series.
   - **Columns H–M**: Computed daily returns and S&P 500 benchmark returns.
   - **Columns R–V (Rows 4–7)**: Individual asset expected returns, variance, standard deviation, and betas.
   - **Columns R–V (Rows 13–17)**: $5 \times 5$ Covariance matrix.
   - **Columns W–X (Rows 13–17)**: Optimal portfolio weights and return contributions.
   - **Cells Q19–R23**: Portfolio variance, standard deviation, beta, and Sharpe metrics.
4. To re-run the optimization:
   - Go to the **Data** tab $\rightarrow$ **Solver**.
   - Note the pre-loaded configuration targeting cell `$R$19` to Min by changing `$W$13:$W$17`.
   - Click **Solve** to re-verify the optimal weights.

---

## License

This project is open-source and distributed under the [MIT License](LICENSE). Feel free to use the data, methodology, and models for academic and professional analysis.
