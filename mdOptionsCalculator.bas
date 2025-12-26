Attribute VB_Name = "mdOptionsCalculator"
Option Explicit

'Purpose: This module handles options trade calculations and conversions

' Constants for options pricing
Private Const OPTION_CONTRACT_MULTIPLIER As Integer = 100 ' Standard options contract multiplier
Private Const DAYS_PER_YEAR As Integer = 365

' Calculate the intrinsic value of an option at a given underlying price
Public Function CalculateIntrinsicValue(ByVal optionType As String, ByVal strike As Double, ByVal underlyingPrice As Double) As Double
    Dim intrinsicValue As Double
    
    If UCase(optionType) = "PUT" Then
        intrinsicValue = Application.WorksheetFunction.Max(0, strike - underlyingPrice)
    ElseIf UCase(optionType) = "CALL" Then
        intrinsicValue = Application.WorksheetFunction.Max(0, underlyingPrice - strike)
    Else
        intrinsicValue = 0
    End If
    
    CalculateIntrinsicValue = intrinsicValue
End Function

' Calculate the PNL for an options trade given the underlying price at exit
' This is a simplified model that assumes:
' - Options are held to expiration or hit target/stop
' - Time decay is approximated
' - Entry price is the premium paid/received
Public Function CalculateOptionsPNL(ByVal oTrade As clsOptionsTrade, ByVal exitUnderlyingPrice As Double, _
                                     Optional ByVal daysToExpiration As Integer = 0, _
                                     Optional ByVal totalDaysToExpiration As Integer = 30) As Double
    
    Dim dblIntrinsicValue As Double
    Dim dblExitPrice As Double
    Dim dblPNL As Double
    Dim dblTimeDecayFactor As Double
    Dim dblTargetPrice As Double
    Dim dblStopPrice As Double
    Dim blnHitTarget As Boolean
    Dim blnHitStop As Boolean
    Dim dblIntrinsicAtEntry As Double
    Dim dblTimeValueAtEntry As Double
    Dim dblEstimatedTimeValue As Double
    Dim dblEffectiveEntryPrice As Double
    
    ' Check if target or stop loss was hit
    blnHitTarget = False
    blnHitStop = False
    
    If oTrade.Target > 0 Then
        If UCase(oTrade.OptionType) = "PUT" Then
            ' For puts, target is hit when underlying goes below target
            If exitUnderlyingPrice <= oTrade.Target Then
                blnHitTarget = True
                exitUnderlyingPrice = oTrade.Target
            End If
        Else ' CALL
            ' For calls, target is hit when underlying goes above target
            If exitUnderlyingPrice >= oTrade.Target Then
                blnHitTarget = True
                exitUnderlyingPrice = oTrade.Target
            End If
        End If
    End If
    
    If oTrade.StopLoss > 0 And Not blnHitTarget Then
        If UCase(oTrade.OptionType) = "PUT" Then
            ' For puts, stop is hit when underlying goes above stop
            If exitUnderlyingPrice >= oTrade.StopLoss Then
                blnHitStop = True
                exitUnderlyingPrice = oTrade.StopLoss
            End If
        Else ' CALL
            ' For calls, stop is hit when underlying goes below stop
            If exitUnderlyingPrice <= oTrade.StopLoss Then
                blnHitStop = True
                exitUnderlyingPrice = oTrade.StopLoss
            End If
        End If
    End If
    
    ' Calculate intrinsic value at exit
    dblIntrinsicValue = CalculateIntrinsicValue(oTrade.OptionType, oTrade.Strike, exitUnderlyingPrice)
    
    ' Apply time decay (simplified linear model)
    ' Time value decays as we approach expiration
    ' This is a simplified approximation - in reality, time decay is non-linear (theta)
    If totalDaysToExpiration > 0 And daysToExpiration >= 0 Then
        dblTimeDecayFactor = daysToExpiration / totalDaysToExpiration
    Else
        dblTimeDecayFactor = 0 ' Assume expired or at expiration
    End If
    
    ' Simplified exit price calculation:
    ' Exit price = Intrinsic value + (Time value * decay factor)
    ' For simplicity, we'll assume the entry price includes both intrinsic and time value
    ' At exit, we use intrinsic value plus remaining time value
    ' This is a rough approximation - a full model would use Black-Scholes or similar
    
    ' Calculate intrinsic value at entry
    dblIntrinsicAtEntry = CalculateIntrinsicValue(oTrade.OptionType, oTrade.Strike, oTrade.CurrentUnderlyingPrice)
    
    If oTrade.EntryPrice > 0 Then
        ' Use provided entry price
        dblTimeValueAtEntry = Application.WorksheetFunction.Max(0, oTrade.EntryPrice - dblIntrinsicAtEntry)
        dblEffectiveEntryPrice = oTrade.EntryPrice
        
        ' Exit price = intrinsic at exit + remaining time value
        dblExitPrice = dblIntrinsicValue + (dblTimeValueAtEntry * dblTimeDecayFactor)
    Else
        ' If no entry price specified, estimate it based on intrinsic value and time value
        ' This is a simplified estimation - in practice, you should input actual option premiums
        ' Estimate time value as a percentage of underlying price (simplified)
        ' Typical at-the-money options have time value of 1-5% of underlying
        dblEstimatedTimeValue = oTrade.CurrentUnderlyingPrice * 0.02 ' 2% default estimate
        
        ' Adjust based on moneyness
        Dim dblMoneyness As Double
        If oTrade.CurrentUnderlyingPrice > 0 Then
            If UCase(oTrade.OptionType) = "PUT" Then
                dblMoneyness = oTrade.Strike / oTrade.CurrentUnderlyingPrice
            Else
                dblMoneyness = oTrade.CurrentUnderlyingPrice / oTrade.Strike
            End If
            ' Out-of-the-money options have less time value
            If dblMoneyness < 0.95 Or dblMoneyness > 1.05 Then
                dblEstimatedTimeValue = dblEstimatedTimeValue * 0.5
            End If
        End If
        
        dblEffectiveEntryPrice = dblIntrinsicAtEntry + dblEstimatedTimeValue
        
        ' Exit price = intrinsic at exit + remaining time value
        dblExitPrice = dblIntrinsicValue + (dblEstimatedTimeValue * dblTimeDecayFactor)
    End If
    
    ' Calculate PNL based on direction
    
    If UCase(oTrade.Direction) = "BUY" Then
        ' Bought option: PNL = (Exit Price - Entry Price) * Quantity * Contract Multiplier
        dblPNL = (dblExitPrice - dblEffectiveEntryPrice) * oTrade.Quantity * OPTION_CONTRACT_MULTIPLIER
    Else ' SELL
        ' Sold option: PNL = (Entry Price - Exit Price) * Quantity * Contract Multiplier
        dblPNL = (dblEffectiveEntryPrice - dblExitPrice) * oTrade.Quantity * OPTION_CONTRACT_MULTIPLIER
    End If
    
    CalculateOptionsPNL = dblPNL
End Function

' Generate a distribution of PNL values from an options trade by simulating underlying price movements
' This creates a Monte Carlo distribution that can be fed into the main simulation
Public Function GenerateOptionsPNLDistribution(ByVal oTrade As clsOptionsTrade, _
                                                ByVal numSimulations As Integer, _
                                                ByVal underlyingVolatility As Double, _
                                                Optional ByVal underlyingDrift As Double = 0) As Variant
    
    Dim arrPNL() As Double
    Dim i As Integer
    Dim dblSimulatedPrice As Double
    Dim dblDaysToExp As Integer
    Dim dblTotalDays As Integer
    Dim dblRandom As Double
    
    ReDim arrPNL(1 To numSimulations)
    
    ' Estimate days to expiration (simplified - assumes current date parsing)
    ' For now, we'll use a default of 30 days if not specified
    dblTotalDays = 30 ' Default
    dblDaysToExp = dblTotalDays ' Assume we're at entry
    
    ' Generate random underlying price movements using normal distribution
    ' Simplified model: Price change follows normal distribution
    For i = 1 To numSimulations
        Randomize (Timer + i)
        
        ' Generate random price movement
        ' Using Box-Muller transform for normal distribution
        Dim dblU1 As Double, dblU2 As Double
        Dim dblZ As Double
        
        dblU1 = Rnd()
        dblU2 = Rnd()
        dblZ = Sqr(-2 * Log(dblU1)) * Cos(2 * Application.WorksheetFunction.Pi() * dblU2)
        
        ' Calculate simulated price: S = S0 * exp((drift - 0.5*vol^2)*t + vol*sqrt(t)*Z)
        ' Simplified: S = S0 * (1 + drift*t + vol*sqrt(t)*Z)
        Dim dblTimeFactor As Double
        dblTimeFactor = dblTotalDays / DAYS_PER_YEAR
        
        Dim dblPriceChange As Double
        dblPriceChange = underlyingDrift * dblTimeFactor + underlyingVolatility * Sqr(dblTimeFactor) * dblZ
        
        dblSimulatedPrice = oTrade.CurrentUnderlyingPrice * (1 + dblPriceChange)
        
        ' Calculate PNL for this simulated price
        arrPNL(i) = CalculateOptionsPNL(oTrade, dblSimulatedPrice, dblDaysToExp, dblTotalDays)
    Next i
    
    GenerateOptionsPNLDistribution = arrPNL
End Function

' Parse a trade string and create an options trade object
' Format examples:
' "buy 200 puts QQQ Exp. 16 of DEC strike 618 Target 617 Stop loss 623"
' "buy 500 puts QQQ Exp. 24 of DEC strike 613, stop loss 618"
Public Function ParseOptionsTradeString(ByVal tradeString As String, _
                                        ByVal currentUnderlyingPrice As Double, _
                                        Optional ByVal entryPrice As Double = 0) As clsOptionsTrade
    
    Dim oTrade As New clsOptionsTrade
    Dim arrWords() As String
    Dim i As Integer
    Dim strWord As String
    Dim blnFoundDirection As Boolean
    Dim blnFoundQuantity As Boolean
    Dim blnFoundType As Boolean
    Dim blnFoundStrike As Boolean
    
    ' Initialize
    oTrade.CurrentUnderlyingPrice = currentUnderlyingPrice
    oTrade.EntryPrice = entryPrice
    
    ' Split the string into words
    tradeString = Trim(tradeString)
    arrWords = Split(tradeString, " ")
    
    ' Parse the string
    i = 0
    Do While i <= UBound(arrWords)
        strWord = UCase(Trim(arrWords(i)))
        
        ' Check for direction
        If (strWord = "BUY" Or strWord = "SELL") And Not blnFoundDirection Then
            oTrade.Direction = strWord
            blnFoundDirection = True
            i = i + 1
            ' Next word should be quantity
            If i <= UBound(arrWords) Then
                oTrade.Quantity = Val(arrWords(i))
                blnFoundQuantity = True
            End If
        ' Check for option type
        ElseIf (strWord = "PUT" Or strWord = "PUTS" Or strWord = "CALL" Or strWord = "CALLS") And Not blnFoundType Then
            If strWord = "PUT" Or strWord = "PUTS" Then
                oTrade.OptionType = "PUT"
            Else
                oTrade.OptionType = "CALL"
            End If
            blnFoundType = True
        ' Check for underlying
        ElseIf strWord = "QQQ" Or strWord = "SPY" Or strWord = "AAPL" Then
            oTrade.Underlying = strWord
        ' Check for strike
        ElseIf strWord = "STRIKE" Then
            i = i + 1
            If i <= UBound(arrWords) Then
                oTrade.Strike = Val(Replace(arrWords(i), "$", ""))
                blnFoundStrike = True
            End If
        ' Check for target
        ElseIf strWord = "TARGET" Then
            i = i + 1
            If i <= UBound(arrWords) Then
                oTrade.Target = Val(Replace(arrWords(i), "$", ""))
            End If
        ' Check for stop loss
        ElseIf strWord = "STOP" Or strWord = "STOPLOSS" Then
            i = i + 1
            If i <= UBound(arrWords) Then
                ' Handle "loss" if present
                If UCase(Trim(arrWords(i))) = "LOSS" Then
                    i = i + 1
                End If
                If i <= UBound(arrWords) Then
                    oTrade.StopLoss = Val(Replace(arrWords(i), "$", ""))
                End If
            End If
        ' Check for expiration (simplified - just store the text)
        ElseIf strWord = "EXP" Or strWord = "EXP." Then
            ' Store expiration info (simplified parsing)
            Dim strExp As String
            strExp = ""
            i = i + 1
            ' Collect expiration details
            Do While i <= UBound(arrWords) And UCase(Trim(arrWords(i))) <> "STRIKE" And UCase(Trim(arrWords(i))) <> "TARGET" And UCase(Trim(arrWords(i))) <> "STOP"
                strExp = strExp & " " & arrWords(i)
                i = i + 1
            Loop
            i = i - 1 ' Back up one
            oTrade.Expiration = Trim(strExp)
        End If
        
        i = i + 1
    Loop
    
    Set ParseOptionsTradeString = oTrade
End Function

