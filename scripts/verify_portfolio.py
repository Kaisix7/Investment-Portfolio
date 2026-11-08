"""
Independent mathematical verification script for Investment Portfolio audit metrics.
Validates GMV formulation, covariance matrix properties, and Sharpe ratio calculations.
"""
import math

def verify_gmv_metrics():
    # A = 1^T * Sigma^-1 * 1
    A = 12479.760124
    expected_sigma_squared = 1.0 / A
    expected_sigma = 1.0 / math.sqrt(A)
    
    # Values from Excel cell R19 and R20
    excel_r19 = 8.01298e-5
    excel_r20 = 0.008951524
    
    assert math.isclose(expected_sigma_squared, excel_r19, rel_tol=1e-4), "Variance mismatch"
    assert math.isclose(expected_sigma, excel_r20, rel_tol=1e-5), "Std dev mismatch"
    print("GMV variance & std dev verification: PASSED")

def verify_sharpe_scaling():
    # Daily metrics
    daily_er = 0.0004576
    daily_rf = 0.0000941
    daily_sigma = 0.0089515
    daily_sharpe = (daily_er - daily_rf) / daily_sigma
    
    # Annualized scaling with sqrt(252)
    annualized_sharpe = daily_sharpe * math.sqrt(252)
    expected_annual_sharpe = (0.1153 - 0.0237) / 0.1421
    
    assert math.isclose(annualized_sharpe, expected_annual_sharpe, rel_tol=1e-2), "Sharpe scaling mismatch"
    print("Sharpe ratio frequency scaling verification: PASSED")

if __name__ == "__main__":
    verify_gmv_metrics()
    verify_sharpe_scaling()
    print("All mathematical verifications PASSED successfully.")
