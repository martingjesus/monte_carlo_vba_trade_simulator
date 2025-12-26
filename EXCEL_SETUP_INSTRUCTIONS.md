# Excel Workbook Setup Instructions

## Overview

The new options calculator code needs to be imported into your existing `monte_carlo_trade_simulator.xlsm` file. The code files in this repository are exported VBA modules that need to be added to the Excel workbook.

## Step 1: Open the Excel Workbook

1. Open `monte_carlo_trade_simulator.xlsm` in Excel
2. Enable macros if prompted
3. Press `Alt + F11` to open the VBA Editor

## Step 2: Import New VBA Modules

You need to import **3 new files** into the VBA project:

### Import clsOptionsTrade.cls (Class Module)

1. In VBA Editor: `File` → `Import File...` (or right-click project → `Import File...`)
2. Select `clsOptionsTrade.cls`
3. This creates a new class module

### Import mdOptionsCalculator.bas (Standard Module)

1. In VBA Editor: `File` → `Import File...`
2. Select `mdOptionsCalculator.bas`
3. This creates a new standard module

### Import mdOptionsProcessor.bas (Standard Module)

1. In VBA Editor: `File` → `Import File...`
2. Select `mdOptionsProcessor.bas`
3. This creates a new standard module

## Step 3: Update Existing Module

The file `mdRun.bas` has been updated. You have two options:

### Option A: Replace the existing module (Recommended)

1. In VBA Editor, find the module `mdRun`
2. Delete it (right-click → `Remove mdRun`)
3. Import the updated `mdRun.bas` file

### Option B: Manual update

1. Open the existing `mdRun` module
2. Replace the `fncGetTrades()` function with the new version that supports options

## Step 4: Set Up Control Sheet Named Ranges

Add these named ranges to your **Control** sheet for options trading:

### Method 1: Using Excel's Name Manager

1. Go to the **Control** sheet
2. Select a cell (e.g., B10) and enter the value `600` (current QQQ price)
3. Select that cell
4. Go to `Formulas` → `Name Manager` → `New`
5. Name: `UNDERLYING_PRICE`
6. Refers to: `=Control!$B$10` (adjust cell reference as needed)
7. Click `OK`

Repeat for:
- **UNDERLYING_VOLATILITY** (e.g., cell B11, value: `0.20`)
- **OPTIONS_SIMULATIONS** (e.g., cell B12, value: `1000`)

### Method 2: Quick Setup Table

Create a table on your Control sheet like this:

| Parameter | Cell | Value | Named Range |
|-----------|------|-------|-------------|
| Current QQQ Price | B10 | 600 | UNDERLYING_PRICE |
| Annual Volatility | B11 | 0.20 | UNDERLYING_VOLATILITY |
| Options Simulations | B12 | 1000 | OPTIONS_SIMULATIONS |

Then create named ranges pointing to these cells.

## Step 5: Test the Setup

1. Go to **InputData** sheet
2. In column A, starting from row 2, enter one of your options trades:
   ```
   buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623
   ```
3. Go to **Control** sheet
4. Click the **Run** button
5. The simulation should run and process the options trade

## Step 6: Verify VBA Project Structure

Your VBA project should now have:

### Class Modules:
- `clsEquityCurve`
- `clsOptionsTrade` ← **NEW**
- `clsResult`
- `clsSimulation`
- `clsTestLogger`
- `INameProvider`
- `ThisWorkbook`

### Standard Modules:
- `mdFactory`
- `mdOptionsCalculator` ← **NEW**
- `mdOptionsProcessor` ← **NEW**
- `mdRun` (updated)
- `TestModule_All`
- `TestModule_clsEquityCurve`
- `TestModule_clsResult`
- `TestModule_clsSimulation`

## Troubleshooting

### Error: "Sub or Function not defined"
- **Solution**: Make sure all three new modules are imported correctly
- Check that `mdOptionsCalculator` and `mdOptionsProcessor` are visible in the VBA project

### Error: "Object variable not set"
- **Solution**: Verify that named ranges exist on the Control sheet
- Check that `UNDERLYING_PRICE`, `UNDERLYING_VOLATILITY`, and `OPTIONS_SIMULATIONS` are defined

### Error: "Type mismatch" or "Invalid procedure call"
- **Solution**: Ensure the InputData sheet has valid trade strings
- Check that the Control sheet named ranges have numeric values

### Options trades not being processed
- **Solution**: Verify the trade string format matches the examples
- Check that the trade includes keywords like "put"/"call" and "strike"

## Quick Reference: Named Ranges Needed

| Named Range | Type | Example Value | Purpose |
|-------------|------|---------------|---------|
| `UNDERLYING_PRICE` | Number | 600 | Current price of QQQ |
| `UNDERLYING_VOLATILITY` | Number | 0.20 | Annual volatility (20%) |
| `OPTIONS_SIMULATIONS` | Integer | 1000 | Simulations per trade |
| `TOTAL_RUNS` | Integer | 2500 | (Existing) Monte Carlo runs |
| `LOT_SIZE` | Integer | 1 | (Existing) Lot size |
| `TRADES_IN_YEAR` | Integer | 100 | (Existing) Trades per year |
| `START_EQUITY` | Number | 100000 | (Existing) Starting equity |
| `MARGIN_LIMIT` | Number | 50000 | (Existing) Margin limit |
| `OUTPUT_START_CELL` | Cell | A20 | (Existing) Output location |
| `OUTPUT` | Range | A20:F50 | (Existing) Output range |

## Next Steps

Once setup is complete:
1. Enter your options trades in the **InputData** sheet
2. Set the parameters on the **Control** sheet
3. Run the simulation as usual
4. The system will automatically detect and process options trades

## Notes

- The existing functionality for simple PNL values remains unchanged
- You can mix options trades and simple PNL values in the same InputData sheet
- The system automatically detects which type each entry is

