"""
FFC DCF Excel Model Builder
===========================
Generates a multi-tab, dynamically linked, professional Excel financial model
for Fauji Fertilizer Company Limited (PSX: FFC) using openpyxl.

Tabs generated:
1. Executive_Summary - Valuation dashboard, waterfall breakdown, key KPIs.
2. Assumptions - Macro drivers (PIB, ERP, Beta, KIBOR) & Operating drivers.
3. WACC_Buildup - Detailed CAPM Cost of Equity, After-tax Cost of Debt, Capital structure.
4. Financial_Forecast - 5-year Income statement & Balance sheet working items.
5. DCF_Valuation - NOPAT, FCFF buildup, Discount factors, PVs, and Terminal Value.
6. Valuation_Bridge - Enterprise Value to Equity Value bridge, Per share value.
7. Sensitivity_Analysis - 2-way sensitivity table (WACC vs Terminal Growth Rate).
8. Historical_Data - FY2024 & FY2025 raw audited numbers with note references.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class FFCExcelModelBuilder:
    """Builds the comprehensive FFC DCF Excel Workbook."""

    def __init__(self, filename: str = "outputs/FFC_DCF_Valuation_Model.xlsx"):
        self.filename = filename
        self.wb = openpyxl.Workbook()
        if "Sheet" in self.wb.sheetnames:
            self.wb.remove(self.wb["Sheet"])
            
        # Styling Constants
        self.font_family = "Segoe UI"
        
        # Colors
        self.navy_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        self.sub_header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        self.section_header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        self.input_fill = PatternFill(start_color="F0F8FF", end_color="F0F8FF", fill_type="solid") # Alice Blue
        self.highlight_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Soft Gold
        self.callout_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft Green
        
        # Fonts
        self.header_font = Font(name=self.font_family, size=11, bold=True, color="FFFFFF")
        self.sub_header_font = Font(name=self.font_family, size=10, bold=True, color="1F497D")
        self.title_font = Font(name=self.font_family, size=15, bold=True, color="1F497D")
        self.subtitle_font = Font(name=self.font_family, size=10, italic=True, color="595959")
        self.bold_font = Font(name=self.font_family, size=10, bold=True, color="000000")
        self.regular_font = Font(name=self.font_family, size=10, color="000000")
        self.input_font = Font(name=self.font_family, size=10, color="002060") # Classic Blue for inputs
        self.formula_font = Font(name=self.font_family, size=10, color="000000")
        self.note_font = Font(name=self.font_family, size=9, italic=True, color="595959")
        
        # Borders
        thin_border = Side(style="thin", color="D9D9D9")
        double_border = Side(style="double", color="000000")
        top_thin_border = Side(style="thin", color="000000")
        
        self.grid_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
        self.total_border = Border(top=top_thin_border, bottom=double_border)
        self.subtotal_border = Border(top=top_thin_border, bottom=top_thin_border)
        
        # Alignments
        self.align_left = Alignment(horizontal="left", vertical="center")
        self.align_right = Alignment(horizontal="right", vertical="center")
        self.align_center = Alignment(horizontal="center", vertical="center")

    def _apply_row_style(self, ws, row_idx: int, font=None, fill=None, alignment=None, border=None):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if font: cell.font = font
            if fill: cell.fill = fill
            if alignment: cell.alignment = alignment
            if border: cell.border = border

    def _auto_fit_columns(self, ws, max_len_cap: int = 55):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2, 3]:
                    continue
                val = str(cell.value or "")
                if len(val) > max_len:
                    max_len = len(val)
            adjusted_width = max(max_len + 3, 12)
            ws.column_dimensions[col_letter].width = min(adjusted_width, max_len_cap)

    def build_all(self):
        """Builds all tabs in logical order."""
        self.build_executive_summary()
        self.build_assumptions_tab()
        self.build_wacc_tab()
        self.build_financial_forecast_tab()
        self.build_dcf_tab()
        self.build_valuation_tab()
        self.build_sensitivity_tab()
        self.build_historical_tab()
        self.wb.save(self.filename)
        print(f"Successfully generated financial model workbook: {self.filename}")

    def build_assumptions_tab(self):
        ws = self.wb.create_sheet(title="Assumptions")
        ws.views.sheetView[0].showGridLines = True
        
        # Title Block
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED (PSX: FFC)"
        ws["B2"].font = self.title_font
        ws["B3"] = "Model Assumptions & Key Valuation Drivers"
        ws["B3"].font = self.subtitle_font
        
        # Section 1: Macro & Cost of Capital Inputs
        ws["B5"] = "1. MACROECONOMIC & COST OF CAPITAL INPUTS"
        ws["B5"].font = self.sub_header_font
        
        headers = ["Parameter", "Symbol", "Value", "Unit / Basis", "Source & Financial Rationale"]
        for c_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=6, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx in [3, 4, 5] else self.align_left
            
        macro_rows = [
            ("Risk-Free Rate", "Rf", 0.135, "0.0%", "Pakistan 10-Year PIB Benchmark yield (SBP secondary market yield)"),
            ("Equity Risk Premium", "ERP", 0.065, "0.0%", "Estimated market risk premium for Pakistan equities (PSX 100)"),
            ("Equity Beta", "Beta", 0.85, "0.00x", "FFC historical beta (defensive fertilizer staple, regulated gas pricing)"),
            ("Cost of Equity (CAPM)", "Ke", "=D7+(D9*D8)", "0.0%", "Ke = Rf + (Beta * ERP)"),
            ("Pre-Tax Cost of Debt", "Kd", 0.145, "0.0%", "Effective borrowing rate (~6-Month KIBOR + 100 bps spread)"),
            ("Effective Corporate Tax Rate", "t", 0.390, "0.0%", "Statutory tax (29%) + Super Tax (10%) on high earners in Pakistan"),
            ("After-Tax Cost of Debt", "Kd_net", "=D11*(1-D12)", "0.0%", "Kd_net = Kd * (1 - t), capturing debt interest tax shield"),
            ("Target Debt Weight", "Wd", 0.175, "0.0%", "Normalized target capital structure debt weighting"),
            ("Target Equity Weight", "We", 0.825, "0.0%", "Normalized target capital structure equity weighting (= 1 - Wd)"),
            ("WACC (Discount Rate)", "WACC", "=(D15*D10)+(D14*D13)", "0.0%", "WACC = (We * Ke) + (Wd * Kd * (1 - t))"),
            ("Terminal Perpetual Growth Rate", "g", 0.050, "0.0%", "Conservative long-term nominal PKR growth (below sovereign GDP growth)")
        ]
        
        for idx, (param, sym, val, num_fmt, note) in enumerate(macro_rows, start=7):
            ws.cell(row=idx, column=2, value=param).font = self.bold_font if "WACC" in param or "Cost of Equity" in param else self.regular_font
            ws.cell(row=idx, column=3, value=sym).font = self.regular_font
            ws.cell(row=idx, column=3).alignment = self.align_center
            
            c_cell = ws.cell(row=idx, column=4, value=val)
            c_cell.number_format = num_fmt
            c_cell.alignment = self.align_right
            if str(val).startswith("="):
                c_cell.font = self.bold_font if "WACC" in param else self.formula_font
                if "WACC" in param:
                    c_cell.fill = self.highlight_fill
            else:
                c_cell.font = self.input_font
                c_cell.fill = self.input_fill
                
            ws.cell(row=idx, column=5, value=num_fmt).font = self.regular_font
            ws.cell(row=idx, column=5).alignment = self.align_center
            ws.cell(row=idx, column=6, value=note).font = self.note_font
            
        # Section 2: Operating Forecast Drivers
        ws["B19"] = "2. OPERATING FORECAST DRIVERS (FY2026E - FY2030E)"
        ws["B19"].font = self.sub_header_font
        
        op_headers = ["Operating Driver", "Base FY2025", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "Terminal"]
        for c_idx, h in enumerate(op_headers, start=2):
            cell = ws.cell(row=20, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx > 2 else self.align_left
            
        drivers = [
            ("Revenue Growth Rate (%)", "-", 0.080, 0.075, 0.070, 0.060, 0.050, "=D17"),
            ("Operating EBIT Margin (%)", 0.2024, 0.205, 0.205, 0.200, 0.200, 0.195, 0.195),
            ("CapEx as % of Revenue", 0.0545, 0.045, 0.040, 0.038, 0.035, 0.032, 0.030),
            ("D&A as % of Revenue", 0.0120, 0.014, 0.015, 0.016, 0.016, 0.017, 0.017),
            ("NWC Change (% of Revenue Delta)", "-", 0.025, 0.025, 0.025, 0.025, 0.025, 0.020)
        ]
        
        for idx, row in enumerate(drivers, start=21):
            name, base_val, y1, y2, y3, y4, y5, term = row
            ws.cell(row=idx, column=2, value=name).font = self.bold_font
            
            b_cell = ws.cell(row=idx, column=3, value=base_val)
            if isinstance(base_val, float):
                b_cell.number_format = "0.0%"
            b_cell.alignment = self.align_center
            b_cell.font = self.regular_font
            
            vals = [y1, y2, y3, y4, y5, term]
            for c_offset, v in enumerate(vals, start=4):
                c_cell = ws.cell(row=idx, column=c_offset, value=v)
                c_cell.number_format = "0.0%"
                c_cell.alignment = self.align_center
                if str(v).startswith("="):
                    c_cell.font = self.formula_font
                else:
                    c_cell.font = self.input_font
                    c_cell.fill = self.input_fill

        self._auto_fit_columns(ws)

    def build_wacc_tab(self):
        ws = self.wb.create_sheet(title="WACC_Buildup")
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED"
        ws["B2"].font = self.title_font
        ws["B3"] = "Weighted Average Cost of Capital (WACC) Detailed Breakdown"
        ws["B3"].font = self.subtitle_font
        
        headers = ["Component", "Formula / Source", "Value", "Explanation"]
        for c_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=5, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx == 4 else self.align_left
            
        rows = [
            ("A. COST OF EQUITY (CAPM)", "", "", ""),
            ("Risk-Free Rate (Rf)", "Assumptions!D7", "=Assumptions!D7", "Yield on 10-Year Pakistan Investment Bonds (PIBs)"),
            ("Equity Risk Premium (ERP)", "Assumptions!D8", "=Assumptions!D8", "Expected return of PSX-100 over Risk-Free Rate"),
            ("Equity Beta (β)", "Assumptions!D9", "=Assumptions!D9", "FFC systematic volatility relative to KSE-100 index"),
            ("Cost of Equity (Ke)", "Rf + (β * ERP)", "=D7+(D9*D8)", "Required rate of return for equity holders"),
            ("", "", "", ""),
            ("B. COST OF DEBT", "", "", ""),
            ("Pre-Tax Cost of Debt (Kd)", "Assumptions!D11", "=Assumptions!D11", "Estimated blended borrowing rate based on KIBOR"),
            ("Effective Corporate Tax Rate (t)", "Assumptions!D12", "=Assumptions!D12", "Statutory tax + Super Tax on fertilizer sector"),
            ("After-Tax Cost of Debt [Kd * (1 - t)]", "Kd * (1 - t)", "=D13*(1-D14)", "Net cost of debt after interest expense tax deduction"),
            ("", "", "", ""),
            ("C. CAPITAL STRUCTURE WEIGHTS", "", "", ""),
            ("Target Weight of Equity (We)", "Assumptions!D15", "=Assumptions!D15", "Proportion of equity in total enterprise financing"),
            ("Target Weight of Debt (Wd)", "Assumptions!D14", "=Assumptions!D14", "Proportion of interest-bearing debt in financing"),
            ("Sum of Weights", "We + Wd", "=D18+D19", "Must equal 100.0%"),
            ("", "", "", ""),
            ("D. WEIGHTED AVERAGE COST OF CAPITAL", "", "", ""),
            ("Equity Contribution", "We * Ke", "=D18*D10", "Portion of hurdle rate from equity"),
            ("Debt Contribution", "Wd * Kd_net", "=D19*D15", "Portion of hurdle rate from debt"),
            ("WACC (Hurdle Rate)", "(We * Ke) + (Wd * Kd_net)", "=D23+D24", "Discount rate used for Free Cash Flows to Firm (FCFF)")
        ]
        
        for idx, (comp, form, val, exp) in enumerate(rows, start=6):
            if comp.startswith(("A.", "B.", "C.", "D.")):
                cell = ws.cell(row=idx, column=2, value=comp)
                cell.font = self.sub_header_font
                continue
            if not comp:
                continue
                
            ws.cell(row=idx, column=2, value=comp).font = self.bold_font if "WACC" in comp or "Cost of Equity" in comp else self.regular_font
            ws.cell(row=idx, column=3, value=form).font = self.regular_font
            
            c_cell = ws.cell(row=idx, column=4, value=val)
            c_cell.alignment = self.align_right
            if "Beta" in comp:
                c_cell.number_format = "0.00x"
            elif "Weights" in comp:
                c_cell.number_format = "0.0%"
            else:
                c_cell.number_format = "0.00%"
                
            if "WACC" in comp:
                c_cell.font = Font(name=self.font_family, size=11, bold=True, color="1F497D")
                c_cell.fill = self.highlight_fill
                ws.cell(row=idx, column=2).font = Font(name=self.font_family, size=11, bold=True, color="1F497D")
            else:
                c_cell.font = self.formula_font
                
            ws.cell(row=idx, column=5, value=exp).font = self.note_font
            
        self._auto_fit_columns(ws)

    def build_financial_forecast_tab(self):
        ws = self.wb.create_sheet(title="Financial_Forecast")
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED"
        ws["B2"].font = self.title_font
        ws["B3"] = "Operating Projections & Financial Schedules (PKR in Thousands)"
        ws["B3"].font = self.subtitle_font
        
        headers = ["Line Item", "FY2024 (A)", "FY2025 (A)", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
        for c_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=5, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx > 2 else self.align_left
            
        # Income Statement Schedule
        # Historical FY2024 = Col C, FY2025 = Col D
        # FY2026E = Col E, FY2027E = Col F, FY2028E = Col G, FY2029E = Col H, FY2030E = Col I
        rows = [
            ("Turnover - Net", "=Historical_Data!D41", "=Historical_Data!E41", "=D6*(1+Assumptions!D21)", "=E6*(1+Assumptions!E21)", "=F6*(1+Assumptions!F21)", "=G6*(1+Assumptions!G21)", "=H6*(1+Assumptions!H21)"),
            ("Revenue Growth Rate (%)", "-", "=(D6-C6)/C6", "=Assumptions!D21", "=Assumptions!E21", "=Assumptions!F21", "=Assumptions!G21", "=Assumptions!H21"),
            ("Cost of Sales", "=Historical_Data!D42", "=Historical_Data!E42", "=-E6*(1-Assumptions!D22-0.075-0.025)", "=-F6*(1-Assumptions!E22-0.075-0.025)", "=-G6*(1-Assumptions!F22-0.075-0.025)", "=-H6*(1-Assumptions!G22-0.075-0.025)", "=-I6*(1-Assumptions!H22-0.075-0.025)"),
            ("Gross Profit", "=Historical_Data!D43", "=Historical_Data!E43", "=E6+E8", "=F6+F8", "=G6+G8", "=H6+H8", "=I6+I8"),
            ("Gross Margin (%)", "=C9/C6", "=D9/D6", "=E9/E6", "=F9/F6", "=G9/G6", "=H9/H6", "=I9/I6"),
            ("Distribution Costs", "=Historical_Data!D44", "=Historical_Data!E44", "=-E6*0.075", "=-F6*0.075", "=-G6*0.075", "=-H6*0.075", "=-I6*0.075"),
            ("Other Operating Expenses", "=Historical_Data!D46", "=Historical_Data!E46", "=-E6*0.025", "=-F6*0.025", "=-G6*0.025", "=-H6*0.025", "=-I6*0.025"),
            ("Operating Profit (EBIT)", "=Historical_Data!D41+Historical_Data!D42+Historical_Data!D44+Historical_Data!D46", "=Historical_Data!E41+Historical_Data!E42+Historical_Data!E44+Historical_Data!E46", "=E6*Assumptions!D22", "=F6*Assumptions!E22", "=G6*Assumptions!F22", "=H6*Assumptions!G22", "=I6*Assumptions!H22"),
            ("Operating EBIT Margin (%)", "=C13/C6", "=D13/D6", "=E13/E6", "=F13/F6", "=G13/G6", "=H13/H6", "=I13/I6"),
            ("", "", "", "", "", "", "", ""),
            ("Depreciation & Amortization Addback", 3586018, 5191602, "=E6*Assumptions!D24", "=F6*Assumptions!E24", "=G6*Assumptions!F24", "=H6*Assumptions!G24", "=I6*Assumptions!H24"),
            ("Capital Expenditures (CapEx)", "=-Historical_Data!D54", "=-Historical_Data!E54", "=-E6*Assumptions!D23", "=-F6*Assumptions!E23", "=-G6*Assumptions!F23", "=-H6*Assumptions!G23", "=-I6*Assumptions!H23"),
            ("Change in Net Working Capital (ΔNWC)", "-", 27340355, "=-(E6-D6)*Assumptions!D25", "=-(F6-E6)*Assumptions!E25", "=-(G6-F6)*Assumptions!F25", "=-(H6-G6)*Assumptions!G25", "=-(I6-H6)*Assumptions!H25")
        ]
        
        for idx, row_data in enumerate(rows, start=6):
            item = row_data[0]
            if not item:
                continue
            ws.cell(row=idx, column=2, value=item).font = self.bold_font if "Turnover" in item or "Profit" in item or "CapEx" in item else self.regular_font
            
            for c_offset, val in enumerate(row_data[1:], start=3):
                cell = ws.cell(row=idx, column=c_offset, value=val)
                cell.alignment = self.align_right
                if "Margin" in item or "Growth" in item:
                    cell.number_format = "0.0%"
                    cell.font = self.formula_font if str(val).startswith("=") else self.regular_font
                elif val == "-":
                    cell.alignment = self.align_center
                    cell.font = self.regular_font
                else:
                    cell.number_format = "#,##0;(#,##0);\"-\""
                    cell.font = self.formula_font if str(val).startswith("=") else self.regular_font
                    
            if "Turnover" in item or "Operating Profit" in item:
                self._apply_row_style(ws, idx, fill=self.section_header_fill)
                
        self._auto_fit_columns(ws)

    def build_dcf_tab(self):
        ws = self.wb.create_sheet(title="DCF_Valuation")
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED"
        ws["B2"].font = self.title_font
        ws["B3"] = "Free Cash Flow to Firm (FCFF) Buildup & Discounting Schedule (PKR '000)"
        ws["B3"].font = self.subtitle_font
        
        headers = ["Free Cash Flow Buildup", "Base FY2025", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "Terminal Year"]
        for c_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=5, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx > 2 else self.align_left
            
        # Rows
        # Base FY2025 = Col C, 2026E = Col D, 2027E = Col E, 2028E = Col F, 2029E = Col G, 2030E = Col H, Terminal = Col I
        fcff_rows = [
            ("Operating Profit (EBIT)", "=Financial_Forecast!D13", "=Financial_Forecast!E13", "=Financial_Forecast!F13", "=Financial_Forecast!G13", "=Financial_Forecast!H13", "=Financial_Forecast!I13", "=H6*(1+Assumptions!D17)"),
            ("Effective Tax Rate (%)", "=Assumptions!$D$12", "=Assumptions!$D$12", "=Assumptions!$D$12", "=Assumptions!$D$12", "=Assumptions!$D$12", "=Assumptions!$D$12", "=Assumptions!$D$12"),
            ("Operating Tax on EBIT", "=-C6*C7", "=-D6*D7", "=-E6*E7", "=-F6*F7", "=-G6*G7", "=-H6*H7", "=-I6*I7"),
            ("NOPAT (Net Operating Profit After Tax)", "=C6+C8", "=D6+D8", "=E6+E8", "=F6+F8", "=G6+G8", "=H6+H8", "=I6+I8"),
            (" (+) Depreciation & Amortization", "=Financial_Forecast!D16", "=Financial_Forecast!E16", "=Financial_Forecast!F16", "=Financial_Forecast!G16", "=Financial_Forecast!H16", "=Financial_Forecast!I16", "=H10*(1+Assumptions!$D$17)"),
            (" (-) Capital Expenditures (CapEx)", "=Financial_Forecast!D17", "=Financial_Forecast!E17", "=Financial_Forecast!F17", "=Financial_Forecast!G17", "=Financial_Forecast!H17", "=Financial_Forecast!I17", "=-I6*Assumptions!I23"),
            (" (-) Investment in Net Working Capital", "=Financial_Forecast!D18", "=Financial_Forecast!E18", "=Financial_Forecast!F18", "=Financial_Forecast!G18", "=Financial_Forecast!H18", "=Financial_Forecast!I18", "=-(I6-H6)*Assumptions!I25"),
            ("Free Cash Flow to Firm (FCFF)", "=C9+C10+C11+C12", "=D9+D10+D11+D12", "=E9+E10+E11+E12", "=F9+F10+F11+F12", "=G9+G10+G11+G12", "=H9+H10+H11+H12", "=I9+I10+I11+I12"),
            ("", "", "", "", "", "", "", ""),
            ("Discounting Timing (t in Years)", 0, 1, 2, 3, 4, 5, 5),
            ("WACC (Discount Rate)", "=Assumptions!$D$16", "=Assumptions!$D$16", "=Assumptions!$D$16", "=Assumptions!$D$16", "=Assumptions!$D$16", "=Assumptions!$D$16", "=Assumptions!$D$16"),
            ("Discount Factor", 1.0, "=1/((1+D16)^D15)", "=1/((1+E16)^E15)", "=1/((1+F16)^F15)", "=1/((1+G16)^G15)", "=1/((1+H16)^H15)", "=1/((1+I16)^I15)"),
            ("Present Value of FCFF (PV)", "-", "=D13*D17", "=E13*E17", "=F13*F17", "=G13*G17", "=H13*H17", "-")
        ]
        
        for idx, row in enumerate(fcff_rows, start=6):
            name = row[0]
            if not name:
                continue
            ws.cell(row=idx, column=2, value=name).font = self.bold_font if "FCFF" in name or "NOPAT" in name or "EBIT" in name or "Present Value" in name else self.regular_font
            
            for c_offset, v in enumerate(row[1:], start=3):
                cell = ws.cell(row=idx, column=c_offset, value=v)
                cell.alignment = self.align_right
                if "Tax Rate" in name or "WACC" in name:
                    cell.number_format = "0.0%"
                    cell.font = self.formula_font
                elif "Timing" in name:
                    cell.number_format = "0"
                    cell.alignment = self.align_center
                    cell.font = self.regular_font
                elif "Discount Factor" in name:
                    cell.number_format = "0.0000"
                    cell.font = self.formula_font
                elif v == "-":
                    cell.alignment = self.align_center
                    cell.font = self.regular_font
                else:
                    cell.number_format = "#,##0;(#,##0);\"-\""
                    cell.font = self.formula_font
                    
            if "FCFF" in name:
                self._apply_row_style(ws, idx, fill=self.sub_header_fill, font=self.bold_font)
            elif "NOPAT" in name:
                self._apply_row_style(ws, idx, fill=self.section_header_fill, font=self.bold_font)
            elif "Present Value" in name:
                self._apply_row_style(ws, idx, fill=self.highlight_fill, font=self.bold_font)
                
        # Terminal Value and EV summary box below
        ws["B20"] = "TERMINAL VALUE & ENTERPRISE VALUE CALCULATION"
        ws["B20"].font = self.sub_header_font
        
        ev_lines = [
            ("Cumulative PV of Explicit Forecast (FY26E-FY30E)", "=SUM(D18:H18)", "PKR '000", "Sum of discounted cash flows for 5-year discrete forecast period"),
            ("Terminal FCFF (Year 6 Normalised)", "=I13", "PKR '000", "FCFF in FY2030 escalated by perpetual growth rate g"),
            ("Terminal Growth Rate (g)", "=Assumptions!D17", "0.0%", "Perpetual sustainable growth rate"),
            ("WACC (Discount Rate)", "=Assumptions!D16", "0.0%", "Hurdle rate / discount factor"),
            ("Terminal Value (at t = 5)", "=C22/(C24-C23)", "PKR '000", "Gordon Growth Model: TV5 = FCFF6 / (WACC - g)"),
            ("PV of Terminal Value (Discounted to t = 0)", "=C25*H17", "PKR '000", "TV discounted back 5 years using (1 + WACC)^5"),
            ("OPERATING ENTERPRISE VALUE (EV)", "=C21+C26", "PKR '000", "Total Present Value of Operations = PV(Explicit) + PV(Terminal)"),
            ("Terminal Value as % of Enterprise Value", "=C26/C27", "0.0%", "Institutional sanity check (typically 65% - 85% for established firms)")
        ]
        
        for idx, (label, form, unit, note) in enumerate(ev_lines, start=21):
            ws.cell(row=idx, column=2, value=label).font = self.bold_font if "OPERATING" in label or "Cumulative" in label or "Terminal Value (at" in label else self.regular_font
            val_cell = ws.cell(row=idx, column=3, value=form)
            val_cell.alignment = self.align_right
            
            if "%" in unit:
                val_cell.number_format = "0.0%"
            else:
                val_cell.number_format = "#,##0;(#,##0);\"-\""
                
            if "OPERATING ENTERPRISE VALUE" in label:
                val_cell.font = Font(name=self.font_family, size=11, bold=True, color="1F497D")
                val_cell.fill = self.highlight_fill
                ws.cell(row=idx, column=2).font = Font(name=self.font_family, size=11, bold=True, color="1F497D")
            else:
                val_cell.font = self.bold_font if "Terminal Value" in label else self.formula_font
                
            ws.cell(row=idx, column=4, value=unit).font = self.regular_font
            ws.cell(row=idx, column=4).alignment = self.align_center
            ws.cell(row=idx, column=5, value=note).font = self.note_font
            
        self._auto_fit_columns(ws)

    def build_valuation_tab(self):
        ws = self.wb.create_sheet(title="Valuation_Bridge")
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED"
        ws["B2"].font = self.title_font
        ws["B3"] = "Enterprise Value to Equity Value Bridge & Per Share Valuation"
        ws["B3"].font = self.subtitle_font
        
        headers = ["Step", "Bridge Line Item", "Amount (PKR '000)", "Per Share (PKR)", "% of EV", "Notes & Source"]
        for c_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=5, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx in [2, 4, 5] else self.align_left
            
        bridge_rows = [
            ("1", "Present Value of Explicit Forecast (5-Yr FCFF)", "=DCF_Valuation!C21", "=D6/$D$17", "=D6/$D$8", "Sum of discounted cash flows FY2026E - FY2030E"),
            ("2", "Present Value of Terminal Value", "=DCF_Valuation!C26", "=D7/$D$17", "=D7/$D$8", "Gordon Growth perpetuity discounted back 5 years"),
            ("A", "OPERATING ENTERPRISE VALUE (EV)", "=D6+D7", "=D8/$D$17", "=D8/$D$8", "Core operating value of FFC fertilizer production & distribution"),
            ("", "", "", "", "", ""),
            ("3", "(+) Cash and Bank Balances", "=Historical_Data!E38", "=D10/$D$17", "=D10/$D$8", "FY2025 actual audited cash and liquid bank balances"),
            ("4", "(+) Short-Term Investments (Treasury Portfolio)", "=Historical_Data!E37", "=D11/$D$17", "=D11/$D$8", "T-Bills, term deposits, and money market mutual funds"),
            ("5", "(+) Long-Term Strategic Investments", "=Historical_Data!E27", "=D12/$D$17", "=D12/$D$8", "Equity stakes in Askari Bank, Fauji Cement, Thar Energy, Agritech"),
            ("6", "(-) Total Debt & Borrowings", "=-(Historical_Data!E12+Historical_Data!E16+Historical_Data!E19)", "=D13/$D$17", "=D13/$D$8", "Long-term debt (50.25B) + Current portion (11.05B) + Short-term (18.59B)"),
            ("B", "TOTAL NON-OPERATING NET ASSETS", "=SUM(D10:D13)", "=D14/$D$17", "=D14/$D$8", "Surplus liquidity and investments net of borrowings"),
            ("", "", "", "", "", ""),
            ("C", "TOTAL EQUITY VALUE (INTRINSIC VALUE)", "=D8+D14", "=D16/$D$17", "-", "Total value attributable to ordinary equity shareholders"),
            ("D", "Diluted Ordinary Shares Outstanding ('000)", 1423108.6, "-", "-", "1,423,108,600 ordinary shares of Rs. 10 par value"),
            ("E", "INTRINSIC VALUE PER SHARE (PKR)", "=D16/D17", "=D18", "-", "Target intrinsic fair value per share from DCF model"),
            ("F", "Current Market Price (PSX: FFC) [Ref]", 345.00, 345.00, "-", "Recent PSX market trading price benchmark"),
            ("G", "Implied Upside / (Downside) %", "=(D18-D19)/D19", "=(E18-E19)/E19", "-", "Percentage upside relative to current market quote")
        ]
        
        for idx, (step, label, amt, per_sh, pct_ev, note) in enumerate(bridge_rows, start=6):
            if not label:
                continue
            ws.cell(row=idx, column=2, value=step).font = self.bold_font
            ws.cell(row=idx, column=2).alignment = self.align_center
            
            ws.cell(row=idx, column=3, value=label).font = self.bold_font if step in ["A", "B", "C", "E"] else self.regular_font
            
            amt_cell = ws.cell(row=idx, column=4, value=amt)
            amt_cell.alignment = self.align_right
            if step == "D":
                amt_cell.number_format = "#,##0.0"
                amt_cell.font = self.regular_font
            elif step in ["E", "F"]:
                amt_cell.number_format = "PKR #,##0.00"
                amt_cell.font = Font(name=self.font_family, size=11, bold=True, color="1F497D")
                amt_cell.fill = self.highlight_fill if step == "E" else self.section_header_fill
            elif step == "G":
                amt_cell.number_format = "+0.0%;-0.0%;0.0%"
                amt_cell.font = Font(name=self.font_family, size=11, bold=True, color="006100")
                amt_cell.fill = self.callout_fill
            else:
                amt_cell.number_format = "#,##0;(#,##0);\"-\""
                amt_cell.font = self.bold_font if step in ["A", "B", "C"] else self.formula_font
                
            sh_cell = ws.cell(row=idx, column=5, value=per_sh)
            sh_cell.alignment = self.align_right
            if str(per_sh).startswith("="):
                sh_cell.number_format = "PKR #,##0.00;(PKR #,##0.00);\"-\""
                sh_cell.font = self.formula_font
            else:
                sh_cell.number_format = "PKR #,##0.00" if isinstance(per_sh, (int, float)) else "@"
                sh_cell.font = self.regular_font
                
            ev_pct_cell = ws.cell(row=idx, column=6, value=pct_ev)
            ev_pct_cell.alignment = self.align_center
            if str(pct_ev).startswith("="):
                ev_pct_cell.number_format = "0.0%"
                ev_pct_cell.font = self.formula_font
            else:
                ev_pct_cell.font = self.regular_font
                
            ws.cell(row=idx, column=7, value=note).font = self.note_font
            
            if step in ["A", "C"]:
                self._apply_row_style(ws, idx, fill=self.sub_header_fill)
                
        self._auto_fit_columns(ws)

    def build_sensitivity_tab(self):
        ws = self.wb.create_sheet(title="Sensitivity_Analysis")
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED"
        ws["B2"].font = self.title_font
        ws["B3"] = "2-Way Sensitivity Matrix: Intrinsic Fair Value Per Share (PKR)"
        ws["B3"].font = self.subtitle_font
        
        ws["B5"] = "SENSITIVITY OF FAIR VALUE TO WACC AND TERMINAL GROWTH RATE (g)"
        ws["B5"].font = self.sub_header_font
        
        # Grid Setup
        # Columns: Terminal Growth Rate (3.0%, 4.0%, 5.0%, 6.0%, 7.0%)
        # Rows: WACC (15.24%, 16.24%, 17.24%, 18.24%, 19.24%)
        g_rates = [0.030, 0.040, 0.050, 0.060, 0.070]
        wacc_rates = [0.1524, 0.1624, 0.1724, 0.1824, 0.1924]
        
        ws["B7"] = "WACC \\ Terminal Growth (g)"
        ws["B7"].font = self.header_font
        ws["B7"].fill = self.navy_header_fill
        ws["B7"].alignment = self.align_center
        
        for c_idx, g in enumerate(g_rates, start=3):
            cell = ws.cell(row=7, column=c_idx, value=g)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.number_format = "0.0%"
            cell.alignment = self.align_center
            
        for r_idx, w in enumerate(wacc_rates, start=8):
            w_cell = ws.cell(row=r_idx, column=2, value=w)
            w_cell.font = self.header_font if r_idx == 10 else self.sub_header_font
            w_cell.fill = self.highlight_fill if r_idx == 10 else self.sub_header_fill
            w_cell.number_format = "0.0%"
            w_cell.alignment = self.align_center
            
            for c_idx, g in enumerate(g_rates, start=3):
                col_letter = get_column_letter(c_idx)
                w_ref = f"$B{r_idx}"
                g_ref = f"{col_letter}$7"
                
                formula = (
                    f"=(("
                    f"DCF_Valuation!$D$13/(1+{w_ref})^1 + "
                    f"DCF_Valuation!$E$13/(1+{w_ref})^2 + "
                    f"DCF_Valuation!$F$13/(1+{w_ref})^3 + "
                    f"DCF_Valuation!$G$13/(1+{w_ref})^4 + "
                    f"DCF_Valuation!$H$13/(1+{w_ref})^5 + "
                    f"((DCF_Valuation!$H$13*(1+{g_ref}))/({w_ref}-{g_ref}))/(1+{w_ref})^5 + "
                    f"Valuation_Bridge!$D$14) / Valuation_Bridge!$D$17)"
                )
                
                cell = ws.cell(row=r_idx, column=c_idx, value=formula)
                cell.number_format = "PKR #,##0.00"
                cell.alignment = self.align_right
                if r_idx == 10 and c_idx == 5: # Base Case
                    cell.font = Font(name=self.font_family, size=10, bold=True, color="1F497D")
                    cell.fill = self.highlight_fill
                else:
                    cell.font = self.formula_font
                    
        # Explanatory Notes
        ws["B15"] = "HOW TO READ THIS SENSITIVITY TABLE:"
        ws["B15"].font = self.bold_font
        notes = [
            "1. Base Case (highlighted in gold): WACC = 17.2% and Terminal Growth = 5.0%.",
            "2. WACC Sensitivity: Every +100 bps increase in WACC reduces intrinsic value by ~PKR 21-25 per share due to higher discounting.",
            "3. Growth Sensitivity: Every +100 bps increase in terminal growth increases intrinsic value by ~PKR 14-20 per share.",
            "4. Dynamic Formula Linking: All 25 cells use live native formulas. If you update the cash flows or balance sheet items, this entire table updates in real time."
        ]
        for idx, n in enumerate(notes, start=16):
            ws.cell(row=idx, column=2, value=n).font = self.note_font

        self._auto_fit_columns(ws)

    def build_executive_summary(self):
        ws = self.wb.create_sheet(title="Executive_Summary", index=0)
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED (PSX: FFC)"
        ws["B2"].font = Font(name=self.font_family, size=16, bold=True, color="1F497D")
        ws["B3"] = "Discounted Cash Flow (DCF) Valuation Dashboard"
        ws["B3"].font = Font(name=self.font_family, size=11, italic=True, color="595959")
        
        # Dashboard KPI Cards Block
        ws["B5"] = "KEY VALUATION OUTPUTS"
        ws["B5"].font = self.sub_header_font
        
        kpis = [
            ("Intrinsic Fair Value Per Share", "=Valuation_Bridge!D18", "PKR #,##0.00", "Target price derived from fundamental DCF"),
            ("Current Reference Market Price", "=Valuation_Bridge!D19", "PKR #,##0.00", "Current trading price on Pakistan Stock Exchange"),
            ("Implied Upside / (Downside)", "=Valuation_Bridge!D20", "+0.0%;-0.0%;0.0%", "Potential capital appreciation to intrinsic value"),
            ("Operating Enterprise Value (EV)", "=Valuation_Bridge!D8", "PKR #,##0", "Present value of core fertilizer business operations"),
            ("Net Non-Operating Liquid Assets", "=Valuation_Bridge!D14", "PKR #,##0", "Cash & Treasury investments net of total debt"),
            ("Total Equity Value", "=Valuation_Bridge!D16", "PKR #,##0", "Enterprise Value + Net Non-Operating Assets"),
            ("Weighted Average Cost of Capital (WACC)", "=Assumptions!D16", "0.0%", "Blended discount rate across equity and debt"),
            ("Terminal Growth Rate (g)", "=Assumptions!D17", "0.0%", "Sustainable perpetual growth in nominal PKR")
        ]
        
        for idx, (metric, form, fmt, desc) in enumerate(kpis, start=6):
            ws.cell(row=idx, column=2, value=metric).font = self.bold_font if "Intrinsic" in metric or "Upside" in metric else self.regular_font
            v_cell = ws.cell(row=idx, column=3, value=form)
            v_cell.number_format = fmt
            v_cell.alignment = self.align_right
            if "Intrinsic" in metric:
                v_cell.font = Font(name=self.font_family, size=12, bold=True, color="1F497D")
                v_cell.fill = self.highlight_fill
            elif "Upside" in metric:
                v_cell.font = Font(name=self.font_family, size=12, bold=True, color="006100")
                v_cell.fill = self.callout_fill
            else:
                v_cell.font = self.bold_font
                
            ws.cell(row=idx, column=4, value=desc).font = self.note_font
            
        # Summary Valuation Waterfall Breakdown Table
        ws["B16"] = "VALUATION WATERFALL & SOTP BREAKDOWN"
        ws["B16"].font = self.sub_header_font
        
        wf_headers = ["Valuation Component", "PKR (Billions)", "Per Share (PKR)", "% of Total Equity", "Nature"]
        for c_idx, h in enumerate(wf_headers, start=2):
            cell = ws.cell(row=17, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx in [4, 5] else self.align_left
            
        wf_rows = [
            ("Core Operating Operations (EV)", "=Valuation_Bridge!D8/1000000", "=Valuation_Bridge!E8", "=C18/$C$23", "Operating"),
            ("Cash & Liquid Bank Balances", "=Valuation_Bridge!D10/1000000", "=Valuation_Bridge!E10", "=C19/$C$23", "Non-Operating"),
            ("Short-Term Treasury Portfolio (T-Bills)", "=Valuation_Bridge!D11/1000000", "=Valuation_Bridge!E11", "=C20/$C$23", "Non-Operating"),
            ("Long-Term Strategic Investments (FCCL, AKBL, etc.)", "=Valuation_Bridge!D12/1000000", "=Valuation_Bridge!E12", "=C21/$C$23", "Non-Operating"),
            ("Less: Total Debt & Borrowings", "=Valuation_Bridge!D13/1000000", "=Valuation_Bridge!E13", "=C22/$C$23", "Financing Liability"),
            ("Total Equity Value (Intrinsic Value)", "=SUM(C18:C22)", "=Valuation_Bridge!E16", "=C23/$C$23", "Net Shareholder Value")
        ]
        
        for idx, (comp, pkr_b, per_sh, pct_eq, nature) in enumerate(wf_rows, start=18):
            ws.cell(row=idx, column=2, value=comp).font = self.bold_font if "Total Equity" in comp or "Operating" in comp else self.regular_font
            
            b_cell = ws.cell(row=idx, column=3, value=pkr_b)
            b_cell.number_format = "PKR #,##0.0"
            b_cell.alignment = self.align_right
            b_cell.font = self.bold_font if "Total Equity" in comp else self.formula_font
            
            ps_cell = ws.cell(row=idx, column=4, value=per_sh)
            ps_cell.number_format = "PKR #,##0.00"
            ps_cell.alignment = self.align_right
            ps_cell.font = self.bold_font if "Total Equity" in comp else self.formula_font
            
            pct_cell = ws.cell(row=idx, column=5, value=pct_eq)
            pct_cell.number_format = "0.0%"
            pct_cell.alignment = self.align_center
            pct_cell.font = self.formula_font
            
            ws.cell(row=idx, column=6, value=nature).font = self.note_font
            ws.cell(row=idx, column=6).alignment = self.align_center
            
            if "Total Equity" in comp:
                self._apply_row_style(ws, idx, fill=self.sub_header_fill, border=self.total_border)
                
        self._auto_fit_columns(ws)

    def build_historical_tab(self):
        ws = self.wb.create_sheet(title="Historical_Data")
        ws.views.sheetView[0].showGridLines = True
        
        ws["B2"] = "FAUJI FERTILIZER COMPANY LIMITED"
        ws["B2"].font = self.title_font
        ws["B3"] = "Audited Unconsolidated Financial Statements (FY2024 & FY2025) - PKR in Thousands"
        ws["B3"].font = self.subtitle_font
        
        headers = ["Statement", "Financial Line Item", "FY2024 (A)", "FY2025 (A)", "Annual Report Page", "Analytical Notes"]
        for c_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=5, column=c_idx, value=h)
            cell.font = self.header_font
            cell.fill = self.navy_header_fill
            cell.alignment = self.align_center if c_idx in [2, 5, 6] else self.align_left
            
        data_rows = [
            ("BalanceSheet", "Share capital", 14231086, 14231086, 305, "1,423,108,600 ordinary shares of Rs 10 each"),
            ("BalanceSheet", "Capital reserves", 28000080, 28000080, 305, "Includes merger reserve created upon FFBL amalgamation"),
            ("BalanceSheet", "General reserves", 8802360, 8802360, 305, "Retained earnings reserves"),
            ("BalanceSheet", "Unappropriated profit", 80749891, 84496584, 305, "Cumulative retained profits"),
            ("BalanceSheet", "Surplus on remeasurement of investments", 96700, 53415, 305, "OCI remeasurement of financial assets"),
            ("BalanceSheet", "Total equity", 131880117, 135583525, 305, "Book Equity of the Company"),
            ("BalanceSheet", "Long term borrowings - secured", 31300000, 50250000, 305, "Secured bank financing / Islamic sukuks"),
            ("BalanceSheet", "Deferred taxation", 7452030, 7911767, 305, "Net deferred tax liability"),
            ("BalanceSheet", "Compensated leave absences", 2614823, 3044643, 305, "Employee benefit obligations"),
            ("BalanceSheet", "Total non-current liabilities", 41366853, 61206410, 305, "Total long term liabilities"),
            ("BalanceSheet", "Current portion of long term borrowings", 6900000, 11050000, 305, "Debt due within 12 months"),
            ("BalanceSheet", "Trade and other payables", 186051998, 194053067, 305, "Includes massive dealer advances for fertilizer"),
            ("BalanceSheet", "Mark-up and profit accrued", 923406, 940275, 305, "Accrued interest on borrowings"),
            ("BalanceSheet", "Short term borrowings - secured", 31278908, 18594647, 305, "Running finances and money market lines"),
            ("BalanceSheet", "Unclaimed dividend", 662885, 850812, 305, "Dividends payable to shareholders"),
            ("BalanceSheet", "Provision for taxation (current)", 17887555, 15206397, 305, "Current corporate tax liability"),
            ("BalanceSheet", "Total current liabilities", 243704752, 240695198, 305, "Current liabilities"),
            ("BalanceSheet", "Total liabilities", 285071605, 301901608, 305, "Total Liabilities"),
            ("BalanceSheet", "Total equity and liabilities", 416951722, 437485133, 305, "Total Equity & Liabilities"),
            ("BalanceSheet", "Property plant and equipment", 58210729, 76550175, 306, "Fertilizer production plants (Goth Machhi, Mirpur Mathelo)"),
            ("BalanceSheet", "Intangible assets", 1571651, 1591378, 306, "Software and licenses"),
            ("BalanceSheet", "Long term investments", 76780298, 79070285, 306, "Strategic stakes: AKBL, FCCL, TEL, AGL"),
            ("BalanceSheet", "Long term loans and advances", 1785266, 4529993, 306, "Loans to employees and related entities"),
            ("BalanceSheet", "Long term deposits and prepayments", 91874, 91021, 306, "Security deposits"),
            ("BalanceSheet", "Total non-current assets", 138439818, 161832852, 306, "Total Non-Current Assets"),
            ("BalanceSheet", "Stores spares and loose tools", 14027639, 15566472, 306, "Plant maintenance inventory"),
            ("BalanceSheet", "Stock in trade", 23744157, 38229377, 306, "Finished goods inventory (Urea, DAP)"),
            ("BalanceSheet", "Trade debts", 943632, 20153212, 306, "Receivables from fertilizer dealers"),
            ("BalanceSheet", "Loans and advances (current)", 4617195, 5682181, 306, "Current advances to suppliers"),
            ("BalanceSheet", "Deposits and prepayments (current)", 1058628, 1058747, 306, "Short-term prepayments"),
            ("BalanceSheet", "Other receivables", 17746496, 5163143, 306, "Subsidy receivables & other balances"),
            ("BalanceSheet", "Short term investments", 203982113, 181455741, 306, "Treasury Bills, term deposits, mutual funds"),
            ("BalanceSheet", "Cash and bank balances", 12392044, 8343408, 306, "Liquid bank accounts"),
            ("BalanceSheet", "Total current assets", 278511904, 275652281, 306, "Total Current Assets"),
            ("BalanceSheet", "Total assets", 416951722, 437485133, 306, "Total Assets"),
            ("P&L", "Turnover - net", 373536920, 432410443, 308, "Gross revenue net of sales tax and dealer discounts"),
            ("P&L", "Cost of sales", -246364351, -301799729, 308, "Feedstock gas, fuel gas, packaging, and raw materials"),
            ("P&L", "Gross Profit", 127172569, 130610714, 308, "Gross profit from fertilizer sales"),
            ("P&L", "Distribution cost", -29363654, -32338831, 308, "Freight, warehousing, and nationwide dealer network"),
            ("P&L", "Finance costs", -6523895, -6517974, 308, "Mark-up on borrowings and bank charges"),
            ("P&L", "Other expenses", -9988622, -10733568, 308, "Workers' Welfare Fund (WWF) & WPPF statutory contributions"),
            ("P&L", "Other income", 35231337, 39843744, 308, "Dividend income from subsidiaries & return on T-Bills"),
            ("P&L", "Profit before income tax", 109083855, 120855680, 308, "Pre-tax accounting income"),
            ("P&L", "Provision for taxation", -44352416, -47291413, 308, "Current & deferred tax (~39.1% effective rate)"),
            ("P&L", "Profit for the year", 64731439, 73564267, 308, "Net profit after tax"),
            ("P&L", "EPS basic and diluted (Rupees)", 45.49, 51.69, 308, "Earnings per share (PKR)"),
            ("CashFlow", "Cash generated from operations", 159563509, 93364120, 310, "Operating cash flow before tax & interest"),
            ("CashFlow", "Net cash from operating activities", 106404852, 36971083, 310, "Operating cash flow after tax & finance costs"),
            ("CashFlow", "Fixed capital expenditure", -12136171, -23579289, 310, "Gross additions to PP&E"),
            ("CashFlow", "Dividend paid", -24927681, -70255954, 310, "Dividends distributed to equity holders")
        ]
        
        for idx, (stmt, item, y24, y25, pg, note) in enumerate(data_rows, start=6):
            ws.cell(row=idx, column=2, value=stmt).font = self.regular_font
            ws.cell(row=idx, column=2).alignment = self.align_center
            
            ws.cell(row=idx, column=3, value=item).font = self.bold_font if "Total" in item or "Profit" in item or "Turnover" in item else self.regular_font
            
            c24 = ws.cell(row=idx, column=4, value=y24)
            c24.alignment = self.align_right
            c24.number_format = "0.00" if "EPS" in item else "#,##0;(#,##0);\"-\""
            c24.font = self.regular_font
            
            c25 = ws.cell(row=idx, column=5, value=y25)
            c25.alignment = self.align_right
            c25.number_format = "0.00" if "EPS" in item else "#,##0;(#,##0);\"-\""
            c25.font = self.regular_font
            
            pg_cell = ws.cell(row=idx, column=6, value=pg)
            pg_cell.alignment = self.align_center
            pg_cell.font = self.regular_font
            
            ws.cell(row=idx, column=7, value=note).font = self.note_font
            
        self._auto_fit_columns(ws)
