# FFC DCF Valuation Model

A comprehensive, institutional-grade Discounted Cash Flow (DCF) valuation of **Fauji Fertilizer Company Limited (PSX: FFC)**, built from scratch with parallel Python and dynamically linked Excel implementations.

---

## Highlights

- **Base Financials:** FY2025 audited unconsolidated financial statements (reflecting full 12-month post-merger operations with FFBL).
- **Core Valuation Methodology:** Free Cash Flow to Firm (FCFF) discounted at Weighted Average Cost of Capital (WACC), with Gordon Growth Terminal Value.
- **Enterprise-to-Equity Bridge:** Full reconciliation accounting for FFC's massive cash, short-term treasury portfolio (T-Bills/mutual funds), long-term strategic investments (AKBL, FCCL, TEL, AGL), and total debt.
- **Valuation Result:**
  - **Intrinsic Fair Value:** **PKR 414.98 per share**
  - **Operating Enterprise Value:** PKR 401.59 Billion (PKR 282.19/share)
  - **Net Non-Operating Liquid Value:** PKR 188.98 Billion (PKR 132.79/share)
  - **Current Reference Market Price (PSX: FFC):** PKR 345.00
  - **Implied Upside:** **+20.3%**
- **Dynamic Excel Model:** `outputs/FFC_DCF_Valuation_Model.xlsx` built with native Excel formulas across 8 cleanly formatted tabs with institutional color-coding and 2-way sensitivity tables.

---

## Directory Structure

```
ffc-dcf/
├── data/
│   ├── processed/
│   │   └── ffc_historical.csv         # Cleaned FY2024-FY2025 audited line items
│   └── raw/
│       ├── Financial-statements-2025.pdf
│       └── sources.md                 # Data sources & page citation references
├── outputs/
│   └── FFC_DCF_Valuation_Model.xlsx   # Live, dynamically linked Excel financial model (gitignored)
├── src/
│   ├── dcf_basics.py                  # Initial discounting prototype
│   ├── dcf_engine.py                  # Core modular DCF engine (WACC, Forecast, TV, Bridge)
│   └── excel_builder.py               # openpyxl workbook generator with native Excel formulas
├── tests/
│   └── test_dcf.py                    # Unit tests for WACC, FCFF, EV bridge, & sensitivity
├── run_valuation.py                   # Runner script: prints report & builds Excel model
├── PROJECT_JOURNAL.md                 # Project development log & key learning insights
└── README.md
```

---

## How to Run

### 1. Run Python Valuation & Rebuild Excel Model
```bash
.venv/bin/python run_valuation.py
```

### 2. Run Test Suite
```bash
.venv/bin/python -m unittest tests/test_dcf.py
```

---

## Key Model Parameters

| Parameter | Value | Source / Rationale |
| :--- | :---: | :--- |
| **Risk-Free Rate ($R_f$)** | 13.50% | 10-Year Pakistan Investment Bond (PIB) secondary yield |
| **Equity Risk Premium (ERP)** | 6.50% | Pakistan market equity risk premium |
| **Equity Beta ($\beta$)** | 0.85x | Defensive essential agricultural staple |
| **Cost of Equity ($K_e$)** | 19.02% | CAPM: $R_f + \beta \times \text{ERP}$ |
| **Pre-Tax Cost of Debt ($K_d$)** | 14.50% | ~6-Month KIBOR + 100 bps spread |
| **Effective Corporate Tax Rate ($t$)** | 39.00% | 29% statutory + 10% super tax |
| **Capital Structure (Equity / Debt)**| 82.5% / 17.5% | Target normalized leverage |
| **WACC (Hurdle Rate)** | **17.24%** | Weighted average discount rate |
| **Terminal Growth Rate ($g$)** | 5.00% | Perpetual nominal PKR growth |
