# Model Guide

## Purpose

This workbook is designed as a source-backed equity valuation model for Eternal Limited.

## Key Tabs

- `01_Cover`: model purpose, version, and high-level outputs
- `02_Dashboard`: current price, target price, recommendation, valuation summary, and data readiness
- `03_Raw_Data`: only source-backed historical values should be entered here
- `04_Historical_FS`: linked historical statements and calculated historical metrics
- `05_Assumptions`: operating drivers, WACC, terminal value, and sensitivity inputs
- `06_Operating_Model`: forecast model and unlevered free cash flow
- `07_Statements`: simplified linked financial statements
- `08_DCF`: intrinsic valuation using Gordon Growth and Exit Multiple methods
- `09_Trading_Comps`: peer trading multiple analysis
- `10_Precedents`: precedent transaction analysis
- `11_Football_Field`: valuation range summary
- `12_Thesis`: analyst-style recommendation draft
- `13_Checks`: model integrity and readiness checks
- `14_Sources_Audit`: source map and audit trail

## Entry Rule

Do not enter historical source values directly into calculation tabs. Enter source values in `03_Raw_Data`; downstream tabs should pull from there.

## Color Convention

- Blue text / yellow fill: editable inputs
- Black text: formulas and calculations
- Green text: links to other workbook sheets

