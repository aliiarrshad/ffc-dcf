"""
FFC DCF Valuation Engine
========================
A modular, rigorous Discounted Cash Flow (DCF) valuation model for
Fauji Fertilizer Company Limited (FFC), listed on the Pakistan Stock Exchange (PSX: FFC).

This engine:
1. Ingests clean historical financial data (FY2025 actuals as base).
2. Computes the Weighted Average Cost of Capital (WACC) using CAPM and local Pakistan benchmarks (PIB, KIBOR).
3. Projects a 5-year forecast of Income Statement items, CapEx, Working Capital, and Free Cash Flow to Firm (FCFF).
4. Calculates Terminal Value using the Gordon Growth Model (Perpetual Growth).
5. Discounts cash flows to compute Enterprise Value (EV).
6. Bridges Enterprise Value to Equity Value by accounting for Cash, Short-Term Investments,
   Long-Term Strategic Investments, and Total Debt.
7. Calculates Intrinsic Fair Value per share.
8. Generates a 2-Way Sensitivity Table (WACC vs. Terminal Growth Rate).
"""

from dataclasses import dataclass, field
from typing import Dict, List, Tuple
import pandas as pd
import numpy as np


@dataclass
class MacroAssumptions:
    """Macroeconomic and Cost of Capital Assumptions (Pakistan Context)."""
    risk_free_rate: float = 0.135       # 10-Year Pakistan Investment Bond (PIB) yield (13.5%)
    equity_risk_premium: float = 0.065   # Equity Risk Premium for Pakistan market (6.5%)
    beta: float = 0.85                  # FFC equity beta (defensive, essential agricultural commodity)
    pre_tax_cost_of_debt: float = 0.145 # Borrowing rate (approx 6M KIBOR + 100 bps spread, 14.5%)
    effective_tax_rate: float = 0.390   # Statutory corporate tax (29%) + Super Tax (10%) = 39.0%
    target_debt_weight: float = 0.175   # Target Capital Structure: Debt Weight (17.5%)
    target_equity_weight: float = 0.825 # Target Capital Structure: Equity Weight (82.5%)
    terminal_growth_rate: float = 0.050 # Sustainable perpetual growth rate in nominal PKR terms (5.0%)

    @property
    def cost_of_equity(self) -> float:
        """Cost of Equity via Capital Asset Pricing Model (CAPM): Ke = Rf + Beta * ERP"""
        return self.risk_free_rate + (self.beta * self.equity_risk_premium)

    @property
    def after_tax_cost_of_debt(self) -> float:
        """After-tax Cost of Debt: Kd_after_tax = Kd * (1 - Tax Rate)"""
        return self.pre_tax_cost_of_debt * (1 - self.effective_tax_rate)

    @property
    def wacc(self) -> float:
        """Weighted Average Cost of Capital (WACC): WACC = (We * Ke) + (Wd * Kd * (1 - t))"""
        return (self.target_equity_weight * self.cost_of_equity) + \
               (self.target_debt_weight * self.after_tax_cost_of_debt)


@dataclass
class OperatingAssumptions:
    """Operating Forecast Drivers (5-Year Horizon)."""
    forecast_years: int = 5
    # Revenue growth trajectory (FY2026E to FY2030E)
    # Reflects steady ~2-3% volumetric ag demand + 5-7% inflation/pricing pass-through
    revenue_growth_rates: List[float] = field(default_factory=lambda: [0.080, 0.075, 0.070, 0.060, 0.050])
    
    # Core operating EBIT margin (FY2025 actual was ~20.24%)
    ebit_margins: List[float] = field(default_factory=lambda: [0.205, 0.205, 0.200, 0.200, 0.195])
    
    # CapEx as % of Revenue (FY2025 had higher turnaround capex; normalizes over time)
    capex_pct_revenue: List[float] = field(default_factory=lambda: [0.045, 0.040, 0.038, 0.035, 0.032])
    
    # Depreciation & Amortization as % of Revenue
    dna_pct_revenue: List[float] = field(default_factory=lambda: [0.014, 0.015, 0.016, 0.016, 0.017])
    
    # Net Working Capital change as % of Revenue increment (fertilizer sales feature low/negative NWC)
    nwc_change_pct_rev_delta: float = 0.025


@dataclass
class HistoricalBase:
    """Base Year (FY2025) Cleaned Financials (in PKR Thousands)."""
    turnover: float = 432_410_443.0          # PKR 432.41B
    cost_of_sales: float = 301_799_729.0     # PKR 301.80B
    gross_profit: float = 130_610_714.0      # PKR 130.61B
    distribution_costs: float = 32_338_831.0 # PKR 32.34B
    other_operating_expenses: float = 10_733_568.0 # PKR 10.73B (WPPF, WWF, admin)
    operating_ebit: float = 87_538_315.0     # Core Operating EBIT = Turnover - COS - Dist - Other Exp
    
    # Cash Flow & Balance Sheet items (FY2025)
    depreciation_amortization: float = 5_191_602.0  # Depreciation (5.188B) + Amortization (3.8M)
    fixed_capital_expenditure: float = 23_579_289.0 # FY2025 PP&E additions
    
    # Non-Operating and Balance Sheet Bridge Items (FY2025)
    cash_and_bank: float = 8_343_408.0              # Cash and bank balances
    short_term_investments: float = 181_455_741.0   # T-Bills, term deposits, mutual funds
    long_term_investments: float = 79_070_285.0     # Strategic equity stakes (AKBL, FCCL, TEL, AGL)
    
    # Total Debt = Long-term borrowings + Current portion + Short-term borrowings
    long_term_borrowings: float = 50_250_000.0
    current_portion_lt_debt: float = 11_050_000.0
    short_term_borrowings: float = 18_594_647.0
    total_debt: float = 79_894_647.0
    
    # Share Statistics
    shares_outstanding: float = 1_423_108.6  # Thousand shares (1,423,108,600 shares)


class FFCDCFEngine:
    """Core DCF Valuation Engine for FFC."""

    def __init__(
        self,
        macro: MacroAssumptions = None,
        operating: OperatingAssumptions = None,
        base: HistoricalBase = None,
        base_year: int = 2025
    ):
        self.macro = macro or MacroAssumptions()
        self.operating = operating or OperatingAssumptions()
        self.base = base or HistoricalBase()
        self.base_year = base_year

    def build_forecast(self) -> pd.DataFrame:
        """
        Projects 5-year operating forecast and Free Cash Flow to Firm (FCFF).
        Returns a pandas DataFrame with all forecast lines.
        """
        years = [self.base_year + i for i in range(1, self.operating.forecast_years + 1)]
        records = []
        
        prev_rev = self.base.turnover
        
        for idx, yr in enumerate(years):
            g = self.operating.revenue_growth_rates[idx]
            rev = prev_rev * (1 + g)
            delta_rev = rev - prev_rev
            
            ebit_margin = self.operating.ebit_margins[idx]
            ebit = rev * ebit_margin
            
            # Tax on Operating Profit (NOPAT = EBIT * (1 - t))
            tax_rate = self.macro.effective_tax_rate
            nopat = ebit * (1 - tax_rate)
            
            # D&A addback
            dna = rev * self.operating.dna_pct_revenue[idx]
            
            # Gross CapEx
            capex = rev * self.operating.capex_pct_revenue[idx]
            
            # Working Capital change (positive means cash outflow to fund working capital)
            delta_nwc = delta_rev * self.operating.nwc_change_pct_rev_delta
            
            # Free Cash Flow to Firm (FCFF)
            # FCFF = NOPAT + D&A - CapEx - Delta NWC
            fcff = nopat + dna - capex - delta_nwc
            
            # Discount period (t = 1, 2, 3, 4, 5)
            t = idx + 1
            # Using end-of-year discounting (or mid-year discounting: (1 + WACC)**(t - 0.5))
            discount_factor = 1.0 / ((1.0 + self.macro.wacc) ** t)
            pv_fcff = fcff * discount_factor
            
            records.append({
                "Year": yr,
                "Period_t": t,
                "Revenue_Growth": g,
                "Turnover": rev,
                "Delta_Revenue": delta_rev,
                "EBIT_Margin": ebit_margin,
                "EBIT": ebit,
                "Tax_Rate": tax_rate,
                "Tax_on_EBIT": ebit * tax_rate,
                "NOPAT": nopat,
                "Depreciation_Amortization": dna,
                "CapEx": capex,
                "Change_in_NWC": delta_nwc,
                "FCFF": fcff,
                "Discount_Factor": discount_factor,
                "PV_of_FCFF": pv_fcff
            })
            
            prev_rev = rev
            
        df = pd.DataFrame(records)
        return df

    def calculate_valuation(self) -> Dict[str, float]:
        """
        Calculates:
        - PV of Explicit Forecast FCFFs
        - Terminal Value & PV of Terminal Value
        - Enterprise Value (EV)
        - EV to Equity Bridge
        - Fair Value per share
        """
        forecast_df = self.build_forecast()
        pv_explicit_fcff = forecast_df["PV_of_FCFF"].sum()
        
        # Terminal Year calculations
        final_fcff = forecast_df["FCFF"].iloc[-1]
        wacc = self.macro.wacc
        g = self.macro.terminal_growth_rate
        
        if wacc <= g:
            raise ValueError(f"WACC ({wacc:.2%}) must be strictly greater than Terminal Growth Rate ({g:.2%})")
            
        # Gordon Growth Model: TV_n = FCFF_n * (1 + g) / (WACC - g)
        terminal_cash_flow = final_fcff * (1 + g)
        terminal_value = terminal_cash_flow / (wacc - g)
        
        # Discount Terminal Value back to base year (t = 5)
        n = self.operating.forecast_years
        tv_discount_factor = 1.0 / ((1.0 + wacc) ** n)
        pv_terminal_value = terminal_value * tv_discount_factor
        
        # Enterprise Value (Operating Value)
        enterprise_value = pv_explicit_fcff + pv_terminal_value
        
        # Equity Value Bridge
        cash_and_st_investments = self.base.cash_and_bank + self.base.short_term_investments
        long_term_investments = self.base.long_term_investments
        total_debt = self.base.total_debt
        
        # Total Equity Value = Enterprise Value + Cash & ST Inv + LT Inv - Total Debt
        net_non_operating_assets = cash_and_st_investments + long_term_investments - total_debt
        equity_value = enterprise_value + net_non_operating_assets
        
        # Per Share Calculations
        # Values are in PKR '000, shares in '000, so (PKR '000 / '000 shares) = PKR per share!
        shares = self.base.shares_outstanding
        value_per_share = equity_value / shares
        operating_value_per_share = enterprise_value / shares
        net_cash_investments_per_share = net_non_operating_assets / shares
        
        return {
            "wacc": wacc,
            "cost_of_equity": self.macro.cost_of_equity,
            "after_tax_cost_of_debt": self.macro.after_tax_cost_of_debt,
            "terminal_growth_rate": g,
            "pv_explicit_fcff": pv_explicit_fcff,
            "terminal_value": terminal_value,
            "pv_terminal_value": pv_terminal_value,
            "enterprise_value": enterprise_value,
            "tv_pct_of_ev": (pv_terminal_value / enterprise_value) * 100.0,
            "cash_and_bank": self.base.cash_and_bank,
            "short_term_investments": self.base.short_term_investments,
            "long_term_investments": self.base.long_term_investments,
            "total_debt": total_debt,
            "net_debt_surplus": total_debt - cash_and_st_investments, # negative means net cash!
            "net_non_operating_assets": net_non_operating_assets,
            "equity_value": equity_value,
            "shares_outstanding": shares,
            "intrinsic_value_per_share": value_per_share,
            "operating_value_per_share": operating_value_per_share,
            "non_operating_value_per_share": net_cash_investments_per_share
        }

    def generate_sensitivity_matrix(
        self,
        wacc_range: List[float] = None,
        growth_range: List[float] = None
    ) -> pd.DataFrame:
        """
        Generates a 2-Way Sensitivity Table:
        Rows: WACC (e.g., 15.0% to 19.0%)
        Columns: Terminal Growth Rate (e.g., 3.0% to 7.0%)
        Cell: Intrinsic Value Per Share (PKR)
        """
        if wacc_range is None:
            base_wacc = self.macro.wacc
            wacc_range = [base_wacc - 0.02, base_wacc - 0.01, base_wacc, base_wacc + 0.01, base_wacc + 0.02]
            
        if growth_range is None:
            base_g = self.macro.terminal_growth_rate
            growth_range = [base_g - 0.02, base_g - 0.01, base_g, base_g + 0.01, base_g + 0.02]
            
        results = {}
        
        original_wacc = self.macro.wacc
        original_g = self.macro.terminal_growth_rate
        
        # We need the explicit forecast FCFF stream
        forecast_df = self.build_forecast()
        explicit_fcffs = forecast_df["FCFF"].tolist()
        final_fcff = explicit_fcffs[-1]
        n = len(explicit_fcffs)
        
        for w in wacc_range:
            row = {}
            for g in growth_range:
                if w <= g:
                    row[f"{g:.1%}"] = np.nan
                    continue
                # PV of explicit cash flows under this WACC
                pv_cf = sum(cf / ((1.0 + w) ** (idx + 1)) for idx, cf in enumerate(explicit_fcffs))
                # Terminal value
                tv = (final_fcff * (1 + g)) / (w - g)
                pv_tv = tv / ((1.0 + w) ** n)
                ev = pv_cf + pv_tv
                
                # Equity value
                eq_val = ev + (self.base.cash_and_bank + self.base.short_term_investments + 
                               self.base.long_term_investments - self.base.total_debt)
                per_share = eq_val / self.base.shares_outstanding
                row[f"{g:.1%}"] = round(per_share, 2)
            results[f"{w:.1%}"] = row
            
        sensitivity_df = pd.DataFrame.from_dict(results, orient="index")
        sensitivity_df.index.name = "WACC \\ Terminal Growth"
        return sensitivity_df
