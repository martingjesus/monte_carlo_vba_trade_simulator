# Options Trading Implementation Summary

## What Was Added

I've extended your Monte Carlo Trade Simulator to support options trades. The system can now:

1. **Parse options trade descriptions** in natural language format
2. **Calculate option PNL** based on underlying price movements, time decay, targets, and stops
3. **Generate PNL distributions** from options trades using Monte Carlo simulation
4. **Integrate seamlessly** with your existing Monte Carlo framework

## New Files Created

### 1. `clsOptionsTrade.cls`
A class to represent an options trade with all parameters:
- Direction (buy/sell)
- Quantity
- Option type (put/call)
- Underlying symbol
- Expiration date
- Strike price
- Target price
- Stop loss price
- Entry price
- Current underlying price

### 2. `mdOptionsCalculator.bas`
Core calculation module with functions:
- `CalculateIntrinsicValue()` - Calculates option intrinsic value
- `CalculateOptionsPNL()` - Calculates PNL for an options trade
- `GenerateOptionsPNLDistribution()` - Creates Monte Carlo distribution of PNL values
- `ParseOptionsTradeString()` - Parses natural language trade descriptions

### 3. `mdOptionsProcessor.bas`
Processing module that:
- Processes multiple options trades
- Converts options trades to PNL distributions
- Detects options trade strings vs. simple PNL values
- Integrates with the existing `fncGetTrades()` function

## Modified Files

### `mdRun.bas`
Updated `fncGetTrades()` function to:
- Detect and process options trades
- Handle both options trades and simple PNL values
- Read parameters from Control sheet (underlying price, volatility, etc.)

## How to Use Your Trades

Enter your trades in the **InputData** sheet, column A:

```
buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623
buy 500 puts QQQ Exp. 24 of DEC strike 613 stop loss 618
buy 500 puts QQQ Exp. 19 of DEC strike 605 Target 602 Stop loss 612
buy 200 calls QQQ Exp. 12 of DEC strike 627 Target 626 Stop loss 621
buy 100 puts QQQ Exp. 10 of DEC strike 623 Target 622 Stop loss 629
```

## Required Setup

Add these named ranges to your **Control** sheet:

1. **UNDERLYING_PRICE** - Current QQQ price (e.g., 600)
2. **UNDERLYING_VOLATILITY** - Annual volatility as decimal (e.g., 0.20 for 20%)
3. **OPTIONS_SIMULATIONS** - Number of simulations per trade (e.g., 1000)

## How It Works

1. **Parsing**: Your trade strings are parsed to extract all parameters
2. **Simulation**: For each trade, the system simulates underlying price movements
3. **PNL Calculation**: For each simulated price, it calculates:
   - Intrinsic value at exit
   - Time decay (simplified linear model)
   - Target/stop loss triggers
   - Final PNL
4. **Distribution**: Creates a PNL distribution from all simulations
5. **Integration**: Feeds the PNL distribution into your existing Monte Carlo framework

## Key Features

✅ **Natural Language Parsing** - Easy to enter trades
✅ **Target & Stop Loss Support** - Automatically handles exit conditions
✅ **Time Decay Modeling** - Simplified but functional
✅ **Mixed Trading** - Can mix options and simple PNL values
✅ **Monte Carlo Integration** - Works with existing simulation framework

## Limitations

⚠️ **Simplified Pricing**: Uses linear time decay, not Black-Scholes
⚠️ **No Greeks**: Delta, gamma, theta, vega not explicitly modeled
⚠️ **Entry Price Estimation**: If not provided, estimates based on simplified model
⚠️ **Single Underlying**: Assumes all trades on same underlying
⚠️ **Constant Volatility**: Uses fixed volatility (no smile/skew)

## Next Steps

1. **Add named ranges** to Control sheet (UNDERLYING_PRICE, UNDERLYING_VOLATILITY, OPTIONS_SIMULATIONS)
2. **Enter your trades** in InputData sheet using the format shown above
3. **Run the simulation** as usual - it will automatically detect and process options trades

## Example Control Sheet Setup

| Cell | Named Range | Value | Description |
|------|-------------|-------|-------------|
| B10 | UNDERLYING_PRICE | 600 | Current QQQ price |
| B11 | UNDERLYING_VOLATILITY | 0.20 | 20% annual volatility |
| B12 | OPTIONS_SIMULATIONS | 1000 | Simulations per trade |

## Testing

The system will work with your existing test framework. You can test options parsing and PNL calculation separately, or integrate with the full Monte Carlo simulation.

## Notes

- Entry prices are estimated if not provided (based on 2% of underlying for ATM options)
- For more accurate results, consider adding actual option premiums
- The system automatically detects options trades vs. simple PNL values
- All existing functionality remains unchanged

