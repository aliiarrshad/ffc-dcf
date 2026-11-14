#!/usr/bin/env python3
"""
FFC DCF Valuation Runner
========================
Executes the discounted cash flow model for Fauji Fertilizer Company Limited (FFC),
prints a comprehensive analytical report to console, and builds the institutional-grade
live Excel financial model.
"""

import sys
from pathlib import Path
from src.dcf_engine import FFCDCFEngine, MacroAssumptions, OperatingAssumptions, HistoricalBase
from src.excel_builder import FFCExcelModelBuilder


def format_pkr_billions(val: float) -> str:
    """Formats PKR in thousands to Billions (PKR B)."""
    return f"PKR {val / 1_000_000:,.2f} B"


def main():
    print("=" * 80)
    print("  FAUJI FERTILIZER COMPANY LIMITED (PSX: FFC) - DCF VALUATION MODEL")
    print("=" * 80)
    
    # 1. Initialize Engine & Run Base Case
    engine = FFCDCFEngine()
    val = engine.calculate_valuation()
    forecast = engine.build_forecast()
    
    print("\n1. MACROECONOMIC ASSUMPTIONS & WACC (Pakistan Context):")
    print("-" * 65)
    print(f"  Risk-Free Rate (10-Yr PIB Benchmark):        {engine.macro.risk_free_rate:6.2%}")
    print(f"  Equity Risk Premium (ERP - PSX 100):          {engine.macro.equity_risk_premium:6.2%}")
    print(f"  FFC Equity Beta:                              {engine.macro.beta:6.2f}x")
    print(f"  Cost of Equity (Ke via CAPM):                 {engine.macro.cost_of_equity:6.2%}")
    print(f"  Pre-Tax Cost of Debt (Kd ~ KIBOR + spread):   {engine.macro.pre_tax_cost_of_debt:6.2%}")
    print(f"  Effective Tax Rate (Corporate 29% + Super):   {engine.macro.effective_tax_rate:6.2%}")
    print(f"  After-Tax Cost of Debt [Kd * (1 - t)]:        {engine.macro.after_tax_cost_of_debt:6.2%}")
    print(f"  Capital Structure (Equity / Debt):            {engine.macro.target_equity_weight:.1%} / {engine.macro.target_debt_weight:.1%}")
    print(f"  Weighted Average Cost of Capital (WACC):      {val['wacc']:6.2%}")
    print(f"  Terminal Perpetual Growth Rate (g):           {val['terminal_growth_rate']:6.2%}")

    print("\n2. 5-YEAR OPERATING FORECAST & FCFF BUILDUP (PKR Millions):")
    print("-" * 80)
    headers = ["Metric", "FY25 (A)", "FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]
    print(f"{headers[0]:<28} " + " ".join(f"{h:>9}" for h in headers[1:]))
    print("-" * 80)
    
    # Base year values
    b_rev = engine.base.turnover / 1_000
    b_ebit = engine.base.operating_ebit / 1_000
    b_nopat = b_ebit * (1 - engine.macro.effective_tax_rate)
    b_dna = engine.base.depreciation_amortization / 1_000
    b_capex = engine.base.fixed_capital_expenditure / 1_000
    
    # Rows
    rows_to_print = [
        ("Revenue (PKR M)", [b_rev] + (forecast["Turnover"] / 1000).tolist()),
        ("Growth Rate (%)", ["-"] + [f"{g:.1%}" for g in forecast["Revenue_Growth"]]),
        ("Operating EBIT (PKR M)", [b_ebit] + (forecast["EBIT"] / 1000).tolist()),
        ("EBIT Margin (%)", [f"{b_ebit/b_rev:.1%}"] + [f"{m:.1%}" for m in forecast["EBIT_Margin"]]),
        ("NOPAT (PKR M)", [b_nopat] + (forecast["NOPAT"] / 1000).tolist()),
        (" (+) D&A (PKR M)", [b_dna] + (forecast["Depreciation_Amortization"] / 1000).tolist()),
        (" (-) CapEx (PKR M)", [b_capex] + (forecast["CapEx"] / 1000).tolist()),
        (" (-) Change in NWC (PKR M)", ["-"] + (forecast["Change_in_NWC"] / 1000).tolist()),
        ("Free Cash Flow (FCFF)", ["-"] + (forecast["FCFF"] / 1000).tolist()),
        ("PV of FCFF (PKR M)", ["-"] + (forecast["PV_of_FCFF"] / 1000).tolist()),
    ]
    
    for label, vals in rows_to_print:
        row_str = f"{label:<28} "
        for v in vals:
            if isinstance(v, (int, float)):
                row_str += f"{v:>9,.0f} "
            else:
                row_str += f"{v:>9} "
        print(row_str)
    print("-" * 80)

    print("\n3. VALUATION SUMMARY & BRIDGE (Enterprise Value to Equity Value):")
    print("-" * 75)
    print(f"  PV of Explicit Forecast (FY26E - FY30E):     {format_pkr_billions(val['pv_explicit_fcff']):>16}")
    print(f"  PV of Terminal Value (at WACC={val['wacc']:.1%}, g={val['terminal_growth_rate']:.1%}):{format_pkr_billions(val['pv_terminal_value']):>16}")
    print(f"  OPERATING ENTERPRISE VALUE (EV):             {format_pkr_billions(val['enterprise_value']):>16}  (100.0%)")
    print(f"    Terminal Value % of Enterprise Value:      {val['tv_pct_of_ev']:>15.1f}%")
    print("  " + "-" * 60)
    print(f"  (+) Cash and Bank Balances:                  {format_pkr_billions(val['cash_and_bank']):>16}")
    print(f"  (+) Short-Term Treasury Portfolio (T-Bills): {format_pkr_billions(val['short_term_investments']):>16}")
    print(f"  (+) Long-Term Strategic Investments:         {format_pkr_billions(val['long_term_investments']):>16}")
    print(f"  (-) Total Debt & Borrowings:                -{format_pkr_billions(val['total_debt']):>16}")
    print(f"  = TOTAL NON-OPERATING NET ASSETS:            {format_pkr_billions(val['net_non_operating_assets']):>16}")
    print("  " + "=" * 60)
    print(f"  TOTAL EQUITY VALUE (INTRINSIC VALUE):        {format_pkr_billions(val['equity_value']):>16}")
    print(f"  Diluted Shares Outstanding:                  {val['shares_outstanding']/1000:>15,.2f} M shares")
    print(f"  INTRINSIC VALUE PER SHARE:                   PKR {val['intrinsic_value_per_share']:>12,.2f}")
    print(f"    - Operating Value Per Share:               PKR {val['operating_value_per_share']:>12,.2f}")
    print(f"    - Net Liquid Investments Per Share:        PKR {val['non_operating_value_per_share']:>12,.2f}")
    print(f"  Current Reference Market Price (PSX: FFC):   PKR {345.00:>12,.2f}")
    upside = (val['intrinsic_value_per_share'] - 345.0) / 345.0
    print(f"  Implied Valuation Upside / (Downside):       {upside:>15.1%}")
    print("=" * 75)

    print("\n4. 2-WAY SENSITIVITY TABLE (Intrinsic Value Per Share in PKR):")
    print("-" * 65)
    sens = engine.generate_sensitivity_matrix()
    print(sens.to_string())
    print("-" * 65)

    # 2. Build the Excel Workbook
    output_dir = Path("outputs")
    output_dir.mkdir(exist_ok=True)
    excel_path = output_dir / "FFC_DCF_Valuation_Model.xlsx"
    print(f"\nBuilding Excel financial model workbook at: {excel_path} ...")
    builder = FFCExcelModelBuilder(str(excel_path))
    builder.build_all()
    print("Done! The financial model has been successfully evaluated and saved.")


if __name__ == "__main__":
    main()
