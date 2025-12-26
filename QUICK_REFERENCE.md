# Quick Reference: Your Options Trades

## Format Your Trades Like This

Copy these into the **InputData** sheet, column A:

```
buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623
buy 500 puts QQQ Exp. 24 of DEC strike 613 stop loss 618
buy 500 puts QQQ Exp. 19 of DEC strike 605 Target 602 Stop loss 612
buy 200 calls QQQ Exp. 12 of DEC strike 627 Target 626 Stop loss 621
buy 100 puts QQQ Exp. 10 of DEC strike 623 Target 622 Stop loss 629
```

## Control Sheet Setup (Required)

Add these named ranges in your **Control** sheet:

| Named Range | Example Value | What It Means |
|-------------|---------------|---------------|
| `UNDERLYING_PRICE` | 600 | Current QQQ price |
| `UNDERLYING_VOLATILITY` | 0.20 | 20% annual volatility |
| `OPTIONS_SIMULATIONS` | 1000 | How many price simulations per trade |

## How It Works

1. ✅ Enter trades in natural language format
2. ✅ System parses and simulates price movements
3. ✅ Calculates PNL for each scenario
4. ✅ Feeds into your existing Monte Carlo simulation

## That's It!

Just run the simulation as usual. The system automatically:
- Detects options trades
- Processes them
- Combines with any simple PNL values
- Runs the Monte Carlo simulation

## Tips

- **Volatility**: Use historical or implied volatility from your options chain
- **Underlying Price**: Update to current market price for accuracy
- **Simulations**: More = better accuracy but slower (1000 is a good default)

