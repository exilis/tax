# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a Japanese tax and social insurance optimization program built for Google Apps Script. The program calculates optimal income allocation across multiple individuals and a corporation (合同会社P&I) to maximize total after-tax wealth (手取り + 税引後利益).

**Language**: The codebase is in Japanese. All variable names, comments, and UI elements use Japanese terminology.

## Key Business Context

### People and Entities
- **林佑樹 (Hayashi Yuki)**: Receives fixed salary from Veltra (¥9.9M/year) + variable compensation from P&I (business income). Individual business owner with ¥6.5M blue-form deduction. Lives in Kobe.
- **土井郁子 (Doi Ikuko)**: Employee or executive. Company-funded corporate DC (¥660k/year). Lives in Nishinomiya.
- **Linh**: Employee living in Kobe.
- **配偶者専従者 (Spouse/Family Employee)**: Receives fixed ¥960k/year salary.
- **合同会社P&I (P&I LLC)**: Consulting company with variable revenue. Pays salaries to employees and compensation to Hayashi.

### Optimization Variables
The program searches for optimal monthly salaries (林報酬, 土井給与, Linh給与) and annual outsourcing fees (事務委託費) to maximize: **Total take-home + P&I after-tax profit**

### Tax Jurisdictions
- Kobe City (神戸市): Hayashi, Linh, spouse
- Nishinomiya City (西宮市): Doi

## How to Use

### Setup and Run
1. Open the Google Spreadsheet containing this script
2. Menu: **税金最適化 → 初期設定** - Creates "optimization" input sheet with default values
3. Adjust input parameters in the "optimization" sheet as needed
4. Menu: **税金最適化 → 最適化実行（手取り+利益の最大化）** - Runs optimization across 21 revenue scenarios
5. Results appear in "最適化結果" sheet with horizontal comparison across revenue variations

### No Testing Framework
This is a standalone Google Apps Script with no unit tests or build process.

## Code Architecture

### Main Entry Points
- `onOpen()` (line 18): Creates custom menu in Google Sheets
- `setupSheet()` (line 29): Initializes input/output sheets
- `runOptimization()` (line 496): Main optimization loop with grid search

### Tax Calculation Functions
- `calcIncomeTax()` (line 148): Progressive income tax with 2.1% reconstruction surtax
- `calcResidentTax()` (line 186): Municipal resident tax (10% + flat ¥5,300)
- `calcCorporateTax()` (line 200): Corporate tax including local taxes (~37% effective rate)
- `calcKojinJigyoTax()` (line 238): Individual business tax (5% after ¥2.9M deduction)

### Social Insurance Functions
- `calcShakaihoken()` (line 117): Annual social insurance (health + pension) based on standard monthly salary
- `getStandardSalary()` (line 124): Maps actual salary to standardized bracket (58k-650k range)

### Core Calculation
- `calcTotalCost()` (line 274): Computes all taxes, insurance, take-home pay for given parameter set. Returns 40+ metrics including:
  - Individual take-home amounts (林、土井、Linh、配偶者)
  - Corporate profit after tax
  - Furusato tax donation limits (ふるさと納税上限額)
  - Effective tax rates

### Optimization Strategy
Grid search across 4 dimensions:
- 事務委託費 (outsourcing fee): ¥0-10M in ¥500k steps
- 3 monthly salaries: ¥150k-800k in ¥50k steps
- Searches ~21,000 combinations per revenue scenario
- Tests 21 revenue scenarios (base + increments of ¥3M up to +¥60M)

**Key optimization parameter:**
- Future retirement bonus tax rate: 27.5% (conservative upper bound)
- This is the theoretical maximum effective tax rate on retirement income
- Derived from: Maximum tax rate 55.95% (45% income tax + surtax + 10% resident tax) × 1/2 (retirement income tax rule) = 27.975%
- Actual rates are 3-15% for realistic retirement amounts (¥30M-¥100M)
- See TAX_COMPARISON.md for detailed mathematical derivation

### Output
- `outputResultsWithVariations()` (line 622): Generates horizontal comparison table with:
  - Optimal allocations (monthly and annual)
  - Maximized total wealth
  - Detailed breakdowns per person/entity
  - Tax rates, insurance costs, furusato donation limits

## Important Constants

### Tax Rates
- Income tax: Progressive 5%-45% (line 148-167)
- Resident tax: Flat 10% + ¥5,300 (line 186-195)
- Corporate tax: 15%/23.2% tiered (line 204-220)
- Individual business tax: 5% (line 243)
- Retirement income tax: (Retirement pay - deduction) × 1/2, then progressive rates apply
  - Deduction: ¥400k/year for first 20 years, then ¥700k/year
  - Effective rate: 0-27.975% (theoretical maximum)

### Social Insurance (Hyogo Prefecture)
- Health insurance (協会けんぽ): 10.29% (line 84)
- Pension (厚生年金): 18.3% (line 85)
- Split 50/50 between employee and employer

### Fixed Deductions
- Basic deduction (基礎控除): ¥480,000
- Blue-form deduction (青色申告控除): ¥650,000
- Business deduction (事業主控除): ¥2,900,000

## Key Japanese Tax Terms
- 給与所得: Salary income
- 事業所得: Business income
- 課税所得: Taxable income
- 所得控除: Income deductions
- 社会保険料: Social insurance premiums
- 標準報酬月額: Standard monthly salary (for insurance calculation)
- ふるさと納税: Hometown tax donation system
- 住宅ローン控除: Housing loan deduction
- 配偶者専従者給与: Spouse/family employee salary
