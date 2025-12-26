Attribute VB_Name = "mdOptionsProcessor"
Option Explicit

'Purpose: This module processes options trades and converts them to PNL distributions for Monte Carlo simulation

' Process a list of options trade strings and generate a combined PNL distribution
' This can be used to feed into the existing Monte Carlo framework
Public Function ProcessOptionsTradesToPNL(ByVal arrTradeStrings As Variant, _
                                          ByVal currentUnderlyingPrice As Double, _
                                          Optional ByVal numSimulationsPerTrade As Integer = 1000, _
                                          Optional ByVal underlyingVolatility As Double = 0.2, _
                                          Optional ByVal underlyingDrift As Double = 0) As Variant
    
    Dim arrAllPNL() As Double
    Dim arrTradePNL As Variant
    Dim oTrade As clsOptionsTrade
    Dim i As Integer, j As Integer
    Dim lTotalPNLCount As Long
    Dim lCurrentIndex As Long
    Dim strTrade As String
    
    ' First pass: count total PNL values
    lTotalPNLCount = 0
    For i = LBound(arrTradeStrings) To UBound(arrTradeStrings)
        strTrade = Trim(CStr(arrTradeStrings(i)))
        If strTrade <> "" Then
            Set oTrade = mdOptionsCalculator.ParseOptionsTradeString(strTrade, currentUnderlyingPrice)
            arrTradePNL = mdOptionsCalculator.GenerateOptionsPNLDistribution(oTrade, numSimulationsPerTrade, underlyingVolatility, underlyingDrift)
            lTotalPNLCount = lTotalPNLCount + UBound(arrTradePNL) - LBound(arrTradePNL) + 1
        End If
    Next i
    
    ' Second pass: combine all PNL values
    ReDim arrAllPNL(1 To lTotalPNLCount)
    lCurrentIndex = 1
    
    For i = LBound(arrTradeStrings) To UBound(arrTradeStrings)
        strTrade = Trim(CStr(arrTradeStrings(i)))
        If strTrade <> "" Then
            Set oTrade = mdOptionsCalculator.ParseOptionsTradeString(strTrade, currentUnderlyingPrice)
            arrTradePNL = mdOptionsCalculator.GenerateOptionsPNLDistribution(oTrade, numSimulationsPerTrade, underlyingVolatility, underlyingDrift)
            
            ' Add this trade's PNL distribution to the combined array
            For j = LBound(arrTradePNL) To UBound(arrTradePNL)
                arrAllPNL(lCurrentIndex) = arrTradePNL(j)
                lCurrentIndex = lCurrentIndex + 1
            Next j
        End If
    Next i
    
    ProcessOptionsTradesToPNL = arrAllPNL
End Function

' Helper function to check if a string represents an options trade
Public Function IsOptionsTradeString(ByVal tradeString As String) As Boolean
    Dim strUpper As String
    strUpper = UCase(Trim(tradeString))
    
    ' Check for keywords that indicate an options trade
    If InStr(strUpper, "PUT") > 0 Or InStr(strUpper, "CALL") > 0 Then
        If InStr(strUpper, "STRIKE") > 0 Or InStr(strUpper, "EXP") > 0 Then
            IsOptionsTradeString = True
            Exit Function
        End If
    End If
    
    IsOptionsTradeString = False
End Function

' Convert options trade strings from a worksheet range to PNL distribution
' This is the main entry point for integrating options trades
Public Function GetOptionsTradesAsPNL(ByVal ws As Worksheet, ByVal rng As Range, _
                                       ByVal currentUnderlyingPrice As Double, _
                                       Optional ByVal numSimulationsPerTrade As Integer = 1000, _
                                       Optional ByVal underlyingVolatility As Double = 0.2) As Variant
    
    Dim arrTradeStrings As Variant
    Dim arrPNL As Variant
    Dim i As Long
    Dim lRow As Long
    Dim collOptionsTrades As New Collection
    Dim strTrade As String
    
    ' Get the range values
    arrTradeStrings = rng.value
    
    ' Process each row
    If IsArray(arrTradeStrings) Then
        ' Multi-cell range
        For i = LBound(arrTradeStrings, 1) To UBound(arrTradeStrings, 1)
            strTrade = Trim(CStr(arrTradeStrings(i, 1)))
            If strTrade <> "" And IsOptionsTradeString(strTrade) Then
                collOptionsTrades.Add strTrade
            End If
        Next i
    Else
        ' Single cell
        strTrade = Trim(CStr(arrTradeStrings))
        If strTrade <> "" And IsOptionsTradeString(strTrade) Then
            collOptionsTrades.Add strTrade
        End If
    End If
    
    ' Convert collection to array
    If collOptionsTrades.Count > 0 Then
        ReDim arrTradeStrings(1 To collOptionsTrades.Count)
        For i = 1 To collOptionsTrades.Count
            arrTradeStrings(i) = collOptionsTrades(i)
        Next i
        
        ' Generate PNL distribution
        arrPNL = ProcessOptionsTradesToPNL(arrTradeStrings, currentUnderlyingPrice, numSimulationsPerTrade, underlyingVolatility)
        GetOptionsTradesAsPNL = arrPNL
    Else
        ' Return empty array
        ReDim arrPNL(1 To 0)
        GetOptionsTradesAsPNL = arrPNL
    End If
    
    Set collOptionsTrades = Nothing
End Function

