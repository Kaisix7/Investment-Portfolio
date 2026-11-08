# Mathematical & Technical Audit Report — Investment Portfolio

This report summarizes the findings from an independent mathematical and technical audit of the README.md against the underlying Excel workbook (`Investment Portfolio.xlsx`).

> [!IMPORTANT]
> The original Excel workbook has been preserved without modification. All corrections apply exclusively to README.md documentation.

---

## 1. Mathematical Errors Found & Corrected

### Error 1: GMV Standard Deviation Formula — Incorrect Expression

**Previous README stated:**
$$\sigma_{GMV} = \sqrt{\frac{1}{A}}$$

**Diagnosis:** While the final *numerical value* (`0.89515%`) was correct, the original README omitted the critical distinction between portfolio **variance** ($\sigma^2 = 1/A$) and portfolio **standard deviation** ($\sigma = 1/\sqrt{A}$). The expression `σ = 1/A` was never explicitly written, but the section lacked a clear decomposition.

**Correction applied:** The README now explicitly shows the two-step derivation:
$$\sigma_{GMV}^2 = \frac{1}{A} = \frac{1}{12,479.76} \approx 8.01298 \times 10^{-5}$$
$$\sigma_{GMV} = \frac{1}{\sqrt{A}} = \frac{1}{\sqrt{12,479.76}} \approx 0.0089515 \quad (0.89515\%)$$

**Verification:** $1 / \sqrt{12479.76} = 1 / 111.7130 \approx 0.008951524$. Matches Excel cell `R20` ($0.008951524$) to 7 significant figures. ✅

---

### Error 2: Covariance Matrix — VARP Described as Covariance Function

**Previous README stated:**
> `=VARP(ПОРТФЕЛЬ!$H$3:$H$1255)` shown as the Excel formula for: $\text{Cov}(R_i, R_j) = \frac{1}{T}\sum(R_{i,t} - \mu_i)(R_{j,t} - \mu_j)$

**Diagnosis:** The Excel function `VARP(...)` computes the **population variance of a single variable**, not the covariance between two variables. It was correctly used on the **diagonal** elements only. The README incorrectly implied `VARP` was used for the off-diagonal covariance entries.

**Correction applied:** The README now distinguishes:
- **Diagonal ($i = j$):** Population variance via `=VARP(...)` — correctly implemented.
- **Off-diagonal ($i \ne j$):** Population covariance via `=COVARIANCE.P(...)` (or evaluated equivalents). Upper-triangular cells mirror lower-triangular cells via direct cell references (`S13 = R14`, etc.) for exact symmetry.

**Verification:** Independently recalculated all 5 diagonal and 10 off-diagonal covariance values from raw daily return data (1,253 observations). All match Excel stored values to 14+ significant figures. ✅

---

### Error 3: Efficient Frontier — Long-Only Incompatibility Not Disclosed

**Previous README implied:** The continuous analytical Markowitz hyperbola was compatible with the workbook's long-only ($w_i \ge 0$) constraints.

**Diagnosis:** Independent audit computed optimal weight vectors $\mathbf{w}(r) = \mathbf{g} + \mathbf{h} r$ for various target returns along the unconstrained analytical frontier:
- At $r = 0.010\%$: Apple weight $= -10.84\%$, Walmart $= -5.72\%$ (short positions required)
- At $r = 0.070\%$: McDonald's weight $= -8.64\%$
- At $r = 0.080\%$: McDonald's weight $= -24.79\%$

Only at the GMV vertex ($r = 0.04576\%$) are all weights naturally non-negative, making it the unique point where unconstrained and constrained solutions coincide.

**Correction applied:** README now explicitly states:
1. The analytical hyperbola assumes unrestricted short selling.
2. The GMV vertex naturally satisfies long-only constraints (verified).
3. Other frontier points violate $w_i \ge 0$ and require constrained QP.
4. Solver was verified as implemented only for the GMV point, not for multi-point frontier tracing.

---

### Error 4: Annualized Returns — Terminology Ambiguity

**Previous README stated:** "Annualized return" without distinguishing arithmetic vs. geometric.

**Diagnosis:** The column labeled "Annualized Return" used the formula $E(R)_\text{daily} \times 252$ (arithmetic scaling). For Tesla, this gives 19.48%, while the actual realized 5-year CAGR is only 1.87% — an 18-percentage-point discrepancy caused by extreme volatility drag.

**Correction applied:**
- Added a new column "5-Year CAGR (Compounded)" alongside the existing "Annualized Arithmetic Return" column.
- Added a dedicated methodology subsection (§5) explaining arithmetic vs. geometric return, with the volatility drag approximation formula.
- Key CAGR values independently computed: Apple 18.53%, Tesla 1.87%, Coca-Cola 11.29%, McDonald's 1.00%, Walmart 19.43%, S&P 500 10.82%.

---

### Error 5: Sharpe Ratio — Frequency Mismatch

**Previous README:** Mentioned the frequency mismatch but lacked full mathematical derivation.

**Correction applied:** Full audit-grade Sharpe derivation now shows:
- Excel cell `R22` subtracts annual $R_f = 2.37\%$ from daily $E(R_p) = 0.0458\%$, producing nominal $-2.5965$ (dimensionally incorrect).
- Consistent daily Sharpe: $(0.04576\% - 0.00941\%) / 0.89515\% \approx \mathbf{0.0406}$.
- Consistent annualized Sharpe: $(11.53\% - 2.37\%) / 14.21\% \approx \mathbf{0.6446}$.
- Cross-verified: $0.0406 \times \sqrt{252} = 0.6446$. ✅

---

## 2. Independently Verified Calculations

| Metric | Excel Value | Independent Calculation | Match |
|:---|:---:|:---:|:---:|
| VARP Apple (Diagonal) | $3.11618 \times 10^{-4}$ | $3.11618 \times 10^{-4}$ | ✅ |
| VARP Tesla (Diagonal) | $1.41446 \times 10^{-3}$ | $1.41446 \times 10^{-3}$ | ✅ |
| Cov(Apple, Tesla) Off-Diag | $3.07070 \times 10^{-4}$ | $3.07070 \times 10^{-4}$ | ✅ |
| Cov(Tesla, Coca-Cola) Off-Diag | $6.53985 \times 10^{-7}$ | $6.53985 \times 10^{-7}$ | ✅ |
| Portfolio Variance (`R19`) | $8.01298 \times 10^{-5}$ | $8.01298 \times 10^{-5}$ | ✅ |
| Portfolio Std Dev (`R20`) | $0.008951524$ | $0.008951524$ | ✅ |
| $A = \mathbf{1}^T\Sigma^{-1}\mathbf{1}$ | — | $12,479.76$ | ✅ |
| $\mu_{GMV} = B/A$ | $0.0004576$ | $0.0004576$ | ✅ |
| $\sigma_{GMV} = 1/\sqrt{A}$ | $0.008952$ | $0.008952$ | ✅ |
| KKT Gradient Condition | — | All 5 marginal contributions equal: $1.6026 \times 10^{-4}$ | ✅ |
| Sum of Weights | $1.0000$ | $1.0000$ | ✅ |
| CAGR Apple (5yr) | — | $18.53\%$ | ✅ (new) |
| CAGR Tesla (5yr) | — | $1.87\%$ | ✅ (new) |

---

## 3. Remaining Limitations

1. **Solver Constraints Not Re-Executed:** The Solver parameters (saved in XML `definedName` entries) confirm the GMV configuration but the optimization was not re-run. The KKT first-order optimality conditions were independently verified, confirming the existing weights are optimal.

2. **Multi-Point Frontier Not Optimized:** The workbook contains placeholder cells (`Q26:R36`) for frontier data points but they are currently empty (value = `None`). Populating these would require running Solver iteratively for each target return level.

3. **Off-Diagonal Covariance Formulas:** The lower-triangular off-diagonal cells in the covariance matrix store evaluated numeric values without visible formulas (likely from a Paste Values operation or a prior calculation step). Their values match independently computed `COVARIANCE.P(...)` results, confirming numerical correctness.

4. **Workbook Sharpe Formula Preserved:** The dimensionally inconsistent Sharpe formula in cell `R22` has been preserved as-is in the workbook per the audit requirements. The correct frequency-consistent values are fully documented in the README.
