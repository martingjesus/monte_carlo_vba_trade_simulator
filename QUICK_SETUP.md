# Quick Setup Guide

## 5-Minute Setup

### 1. Import VBA Files (3 files)

Open Excel → `Alt+F11` → Import these files:
- ✅ `clsOptionsTrade.cls` (Class Module)
- ✅ `mdOptionsCalculator.bas` (Standard Module)  
- ✅ `mdOptionsProcessor.bas` (Standard Module)
- ✅ `mdRun.bas` (Replace existing or update manually)

### 2. Add Named Ranges to Control Sheet

On the **Control** sheet, create these named ranges:

| Name | Example Cell | Example Value |
|------|--------------|---------------|
| `UNDERLYING_PRICE` | B10 | 600 |
| `UNDERLYING_VOLATILITY` | B11 | 0.20 |
| `OPTIONS_SIMULATIONS` | B12 | 1000 |

**How to create:**
1. Enter value in cell (e.g., B10 = 600)
2. Select the cell
3. Go to `Formulas` → `Define Name`
4. Enter the name (e.g., `UNDERLYING_PRICE`)
5. Click OK

### 3. Enter Your Trades

On the **InputData** sheet, column A, row 2 onwards:

```
buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623
buy 500 puts QQQ Exp. 24 of DEC strike 613 stop loss 618
buy 500 puts QQQ Exp. 19 of DEC strike 605 Target 602 Stop loss 612
buy 200 calls QQQ Exp. 12 of DEC strike 627 Target 626 Stop loss 621
buy 100 puts QQQ Exp. 10 of DEC strike 623 Target 622 Stop loss 629
```

### 4. Run Simulation

Go to **Control** sheet → Click **Run** button

**Done!** 🎉

The system will:
- ✅ Detect your options trades
- ✅ Simulate price movements
- ✅ Calculate PNL distributions
- ✅ Run Monte Carlo simulation
- ✅ Show results as usual

