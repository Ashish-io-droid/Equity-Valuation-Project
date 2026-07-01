# Equity Valuation Project - Eternal Limited / Zomato

This repository contains a source-backed equity valuation project for Eternal Limited, formerly Zomato. The project is built around an investment-banking style workflow: collect primary filings, extract historical and quarterly operating data, build a linked financial model, triangulate valuation through multiple methods, and present the final investment view clearly.

The goal is to create a portfolio-quality valuation package that demonstrates financial statement analysis, operating-driver thinking, DCF methodology, comparable-company analysis, precedent-transaction analysis, and analyst-style communication.

## Project Scope

The project is structured to support a full public-company valuation workflow:

- Source collection from annual reports, quarterly filings, investor materials, and research references
- Historical financial statement extraction and validation
- Quarterly operating and financial KPI dataset
- Three-statement model architecture
- Operating assumptions and segment-level drivers
- Discounted cash flow valuation
- Trading comparable company analysis
- Precedent transaction analysis
- Football-field valuation output
- One-page investment thesis and recommendation
- Source audit trail and model checks

## Repository Structure

```text
Eternal-Valuation/
  data/
    annual_reports pdf/          Annual report source PDFs
    quarterly_results pdf/       Quarterly result source PDFs
    extracted_text/              Extracted source text and page references
    master_dataset/              Clean source-backed datasets
    research_package/            Supporting research package extracts
  docs/
    Methodology.md               Modeling and valuation methodology
    Model_Guide.md               Workbook structure and usage notes
    Source_List.md               Source inventory
  excel/
    Eternal_Valuation_Model.xlsx Main valuation model workbook
  outputs/
    v1.1_professional_structure/ Professional workbook structure output
    v2.0_quarterly_data/         Quarterly data pack output
  presentation/
  reports/

valuation_model/
  build_eternal_model.mjs        Main workbook generation script
  build_quarterly_data_model.mjs Quarterly data pack generation script
```

## Key Artifacts

### Main Valuation Model

The main Excel model is organized into a professional valuation workbook structure with tabs for:

- Cover and dashboard
- Raw data
- Historical financial statements
- Assumptions
- Operating model
- Three financial statements
- DCF valuation
- Trading comps
- Precedents
- Football field
- Thesis
- Checks
- Sources audit

### Quarterly Data Pack

The quarterly data pack contains a source-backed dataset for Eternal's quarterly operating and financial metrics. It includes:

- Master extraction table
- Segment KPI view
- Formula-driven checks
- Source audit sheet
- CSV support files

The pack captures metrics such as B2C NOV, adjusted revenue, adjusted EBITDA, food delivery KPIs, quick commerce KPIs, Going-out revenue/NOV, Hyperpure revenue, cash bridge items, and adjusted EBITDA reconciliation items.

### Source Documentation

The repository preserves the source documents used for extraction and validation, including annual reports and quarterly result PDFs. The documentation files record source logic, methodology, model architecture, and audit conventions.

## Valuation Methodology

The model is designed around three valuation perspectives:

1. Discounted Cash Flow

   Projects unlevered free cash flow using operating assumptions, margin forecasts, reinvestment needs, working capital movement, and terminal value methodology.

2. Trading Comparables

   Benchmarks Eternal against relevant listed peers using market valuation multiples such as EV/Revenue, EV/EBITDA, and P/E where applicable.

3. Precedent Transactions

   Uses relevant M&A transactions as a market-based cross-check where deal data is available and comparable.

The outputs are intended to be summarized in a football-field chart and supported by a concise investment thesis.

## Source and Audit Standards

This project follows a primary-source-first approach:

- Annual reports and quarterly filings are preferred as source of truth
- Investor presentations and shareholder letters are used for KPI and segment disclosures
- Secondary sources are used only for cross-checking or contextual support
- Every extracted value is expected to carry source reference, unit, confidence level, and verification status
- Checks are included to validate totals, reconciliations, cash movement, and formula integrity

## How to Use

Open the workbook files in Excel from the `Eternal-Valuation/excel/` and `Eternal-Valuation/outputs/` folders. The raw-data and master-dataset files are designed to make the model auditable: source values sit separately from calculations, and checks show whether reconciliations tie.

To regenerate the workbooks, use the builder scripts in `valuation_model/` with the bundled Node.js environment used for this project.

## Disclaimer

This repository is for educational and portfolio purposes only. It is not investment advice, a recommendation to buy or sell securities, or a substitute for professional financial analysis.
