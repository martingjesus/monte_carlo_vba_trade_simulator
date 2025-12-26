# Options Trading Guide for Monte Carlo Simulator

## Overview

The Monte Carlo Trade Simulator now supports options trades in addition to simple PNL values. This allows you to model options strategies with targets, stop losses, and expiration dates.

## How It Works

The system converts your options trades into a distribution of PNL values by:
1. Parsing your options trade descriptions
2. Simulating underlying price movements using Monte Carlo
3. Calculating option payoffs based on intrinsic value, time decay, targets, and stops
4. Generating a PNL distribution that feeds into the main Monte Carlo simulation

## Input Format

Enter your options trades in the **InputData** sheet, column A, using natural language format:

### Basic Format
```
buy [quantity] [puts/calls] [underlying] Exp. [date] strike [price] Target [price] Stop loss [price]
```

### Examples from Your Trades

1. **Buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623**
   - Buys 200 QQQ put options
   - Expiration: December 16
   - Strike: $618
   - Target: $617 (underlying price)
   - Stop loss: $623 (underlying price)

2. **buy 500 puts QQQ Exp. 24 of DEC strike 613, stop loss 618**
   - Buys 500 QQQ put options
   - Expiration: December 24
   - Strike: $613
   - Stop loss: $618
   - No target specified

3. **buy 200 calls QQQ Exp. 12 of DEC strike 627$ Target 626$ Stop loss 621$**
   - Buys 200 QQQ call options
   - Expiration: December 12
   - Strike: $627
   - Target: $626
   - Stop loss: $621

### Format Notes

- **Direction**: "buy" or "sell" (case insensitive)
- **Quantity**: Number of contracts
- **Option Type**: "puts" or "calls" (case insensitive)
- **Underlying**: Stock/ETF symbol (e.g., "QQQ", "SPY")
- **Expiration**: "Exp." or "Exp" followed by date description
- **Strike**: "strike" followed by price (with or without $)
- **Target**: "Target" followed by underlying price
- **Stop Loss**: "Stop loss" or "Stop" followed by underlying price

## Control Sheet Parameters

Add these named ranges to your **Control** sheet for options trading:

### Required Parameters

1. **UNDERLYING_PRICE** (Cell reference)
   - Current price of the underlying asset (e.g., QQQ)
   - Example: If QQQ is at $600, set this to 600
   - **Default**: 600 if not specified

2. **UNDERLYING_VOLATILITY** (Cell reference)
   - Annual volatility of the underlying (as decimal)
   - Example: 0.20 for 20% annual volatility
   - **Default**: 0.20 (20%) if not specified

3. **OPTIONS_SIMULATIONS** (Cell reference)
   - Number of price simulations per options trade
   - Higher values = more accurate but slower
   - **Default**: 1000 if not specified

### Example Control Sheet Setup

| Parameter Name | Cell | Value | Description |
|---------------|------|-------|-------------|
| UNDERLYING_PRICE | B10 | 600 | Current QQQ price |
| UNDERLYING_VOLATILITY | B11 | 0.20 | 20% annual volatility |
| OPTIONS_SIMULATIONS | B12 | 1000 | Simulations per trade |

## How PNL is Calculated

### For PUT Options (Buy)
- **Intrinsic Value** = Max(0, Strike - Underlying Price)
- **Exit Price** = Intrinsic Value + Remaining Time Value
- **PNL** = (Exit Price - Entry Price) × Quantity × 100

### For CALL Options (Buy)
- **Intrinsic Value** = Max(0, Underlying Price - Strike)
- **Exit Price** = Intrinsic Value + Remaining Time Value
- **PNL** = (Exit Price - Entry Price) × Quantity × 100

### Target and Stop Loss Logic

- **PUT Target**: Hit when underlying price ≤ target price
- **PUT Stop**: Hit when underlying price ≥ stop price
- **CALL Target**: Hit when underlying price ≥ target price
- **CALL Stop**: Hit when underlying price ≤ stop price

### Time Decay

The model uses a simplified linear time decay approximation:
- Time value decays proportionally as expiration approaches
- Full time value at entry, zero at expiration

## Entry Price Estimation

If you don't specify entry prices, the system estimates them based on:
- Current intrinsic value
- Estimated time value (simplified model)

**Note**: For more accurate results, you can manually calculate and input entry prices by:
1. Looking up the option premium in your trading platform
2. Adding a column in InputData for entry prices (future enhancement)

## Mixing Options and Simple PNL

You can mix options trades and simple PNL values in the same InputData sheet:
- Options trades: Use the natural language format
- Simple PNL: Enter numeric values directly

The system automatically detects and processes both types.

## Example: Your Trades

Based on your provided trades, here's how to enter them:

```
buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623
buy 500 puts QQQ Exp. 24 of DEC strike 613 stop loss 618
buy 500 puts QQQ Exp. 19 of DEC strike 605 Target 602 Stop loss 612
buy 200 calls QQQ Exp. 12 of DEC strike 627 Target 626 Stop loss 621
buy 100 puts QQQ Exp. 10 of DEC strike 623 Target 622 Stop loss 629
```

## Limitations and Assumptions

1. **Simplified Pricing Model**: Uses linear time decay, not Black-Scholes
2. **No Greeks**: Delta, gamma, theta, vega are not explicitly modeled
3. **No Early Exercise**: Assumes options are held to expiration or target/stop
4. **Constant Volatility**: Uses fixed volatility (no volatility smile)
5. **No Dividends**: Dividend effects are not included
6. **Single Underlying**: All trades assumed to be on the same underlying

## Tips for Best Results

1. **Set Realistic Volatility**: Use historical volatility or implied volatility from options chain
2. **Use Appropriate Entry Prices**: If possible, input actual option premiums
3. **Adjust Simulation Count**: More simulations = better accuracy but slower
4. **Validate Results**: Compare simulated results with actual historical performance

## Troubleshooting

**Issue**: Options trades not being recognized
- **Solution**: Ensure format includes "put"/"call" and "strike" keywords

**Issue**: PNL seems incorrect
- **Solution**: Check UNDERLYING_PRICE matches current market price
- **Solution**: Verify strike prices and targets/stops are correct

**Issue**: Simulation is slow
- **Solution**: Reduce OPTIONS_SIMULATIONS (try 500 instead of 1000)

## Future Enhancements

Potential improvements:
- Black-Scholes pricing model
- Greeks calculation (delta, gamma, theta, vega)
- Volatility smile/skew
- Early exercise modeling
- Multiple underlyings
- Spread strategies (straddles, strangles, etc.)

