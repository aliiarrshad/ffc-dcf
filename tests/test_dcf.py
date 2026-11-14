"""
Unit Tests for FFC DCF Valuation Model
"""

import unittest
from pathlib import Path
import openpyxl
from src.dcf_engine import FFCDCFEngine, MacroAssumptions, OperatingAssumptions, HistoricalBase
from src.excel_builder import FFCExcelModelBuilder


class TestFFCDCF(unittest.TestCase):

    def setUp(self):
        self.engine = FFCDCFEngine()

    def test_wacc_calculation(self):
        macro = self.engine.macro
        # Ke = 0.135 + (0.85 * 0.065) = 0.135 + 0.05525 = 0.19025
        self.assertAlmostEqual(macro.cost_of_equity, 0.19025, places=5)
        
        # Kd_net = 0.145 * (1 - 0.39) = 0.08845
        self.assertAlmostEqual(macro.after_tax_cost_of_debt, 0.08845, places=5)
        
        # WACC = (0.825 * 0.19025) + (0.175 * 0.08845) = 0.15695625 + 0.01547875 = 0.172435
        expected_wacc = (0.825 * 0.19025) + (0.175 * 0.08845)
        self.assertAlmostEqual(macro.wacc, expected_wacc, places=5)

    def test_forecast_schedule(self):
        df = self.engine.build_forecast()
        self.assertEqual(len(df), 5)
        self.assertEqual(list(df["Year"]), [2026, 2027, 2028, 2029, 2030])
        
        # Revenue should grow monotonically
        revs = df["Turnover"].tolist()
        self.assertTrue(all(revs[i] < revs[i+1] for i in range(len(revs)-1)))
        
        # FCFF = NOPAT + D&A - CapEx - Delta NWC
        for idx, row in df.iterrows():
            expected_fcff = row["NOPAT"] + row["Depreciation_Amortization"] - row["CapEx"] - row["Change_in_NWC"]
            self.assertAlmostEqual(row["FCFF"], expected_fcff, places=2)

    def test_valuation_bridge(self):
        val = self.engine.calculate_valuation()
        
        # Enterprise value should equal PV(explicit) + PV(terminal)
        self.assertAlmostEqual(val["enterprise_value"], val["pv_explicit_fcff"] + val["pv_terminal_value"], places=2)
        
        # Net non-operating assets should equal Cash + ST Inv + LT Inv - Total Debt
        expected_net_non_op = (val["cash_and_bank"] + val["short_term_investments"] + 
                               val["long_term_investments"] - val["total_debt"])
        self.assertAlmostEqual(val["net_non_operating_assets"], expected_net_non_op, places=2)
        
        # Equity value = EV + Net Non Operating Assets
        self.assertAlmostEqual(val["equity_value"], val["enterprise_value"] + val["net_non_operating_assets"], places=2)
        
        # Per share value
        expected_per_share = val["equity_value"] / val["shares_outstanding"]
        self.assertAlmostEqual(val["intrinsic_value_per_share"], expected_per_share, places=2)

    def test_sensitivity_table(self):
        sens = self.engine.generate_sensitivity_matrix()
        self.assertEqual(sens.shape, (5, 5))
        
        # Higher WACC should decrease per share value
        col_5pct = sens["5.0%"].tolist()
        self.assertTrue(all(col_5pct[i] > col_5pct[i+1] for i in range(len(col_5pct)-1)))
        
        # Higher growth rate should increase per share value
        row_base = sens.iloc[2].tolist() # middle row is base WACC
        self.assertTrue(all(row_base[i] < row_base[i+1] for i in range(len(row_base)-1)))

    def test_excel_file_generation(self):
        output_file = "outputs/test_model.xlsx"
        builder = FFCExcelModelBuilder(output_file)
        builder.build_all()
        
        path = Path(output_file)
        self.assertTrue(path.exists())
        self.assertGreater(path.stat().st_size, 5000)
        
        wb = openpyxl.load_workbook(output_file)
        expected_sheets = [
            'Executive_Summary', 'Assumptions', 'WACC_Buildup', 
            'Financial_Forecast', 'DCF_Valuation', 'Valuation_Bridge', 
            'Sensitivity_Analysis', 'Historical_Data'
        ]
        self.assertEqual(wb.sheetnames, expected_sheets)
        path.unlink() # Cleanup test file


if __name__ == "__main__":
    unittest.main()
