Attribute VB_Name = "mdRun"
Option Explicit


Public Sub StartMonteCarloSimulation()
'*******************************************************************************************************

'Purpose: This is the main routine that takes the parameters from the worksheet and runs the simulation

'*******************************************************************************************************

    
    Dim vntTradeList As Variant
    Dim iCalc As Integer
    Dim blnScreenUpdating As Boolean
    Dim collFinalResults As Collection
    Dim oResult As clsResult
    Dim lRow As Long
    Dim iCol As Integer
    Dim oSimulation As clsSimulation
    Dim ws As Worksheet
    Dim intTotalRuns As Integer
    Dim intLotSize As Integer
    Dim dblStartEquity As Double
    Dim dblMarginLimit As Double
    Dim intTradesInYear As Integer
    
    iCalc = Application.Calculation
    Application.Calculation = xlManual
    
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' clear the output from previous results
    Call ClearUI
    
    ' get the list of trades from the input worksheet
    vntTradeList = fncGetTrades()
    If UBound(vntTradeList) = 0 Then
        MsgBox "No trade list found!", vbExclamation + vbOKOnly, "Input Data"
        GoTo Exit_Here
    End If
    
    Set ws = ThisWorkbook.Sheets("Control")
    With ws
    
        ' parameters for the simulation
        intTotalRuns = .Range("TOTAL_RUNS").value
        intLotSize = .Range("LOT_SIZE").value
        intTradesInYear = .Range("TRADES_IN_YEAR").value
        dblStartEquity = .Range("START_EQUITY").value
        dblMarginLimit = .Range("MARGIN_LIMIT").value
        
        'create a simulation object to run with the parameters
        Set oSimulation = mdFactory.CreateSimulation(totalRuns:=intTotalRuns, _
            tradesInYear:=intTradesInYear, lotSize:=intLotSize, TradeList:=vntTradeList, _
            startEquity:=dblStartEquity, margin:=dblMarginLimit)
    
        If Not oSimulation Is Nothing Then
            lRow = .Range("OUTPUT_START_CELL").Row
            iCol = .Range("OUTPUT_START_CELL").Column
            
            'run the simulation
            Set collFinalResults = oSimulation.fncRunProcess()
            
            'output the results of the simulation
            If Not collFinalResults Is Nothing Then
                For Each oResult In collFinalResults
                   
                   .Cells(lRow, iCol).value = oResult.equity
                   .Cells(lRow, iCol + 1).value = oResult.Ruin
                   .Cells(lRow, iCol + 2).value = oResult.MedianDrawdown
                   .Cells(lRow, iCol + 3).value = oResult.MedianProfit
                   .Cells(lRow, iCol + 4).value = oResult.MedianReturn
                   .Cells(lRow, iCol + 5).value = oResult.MedianReturnDD
    
                   lRow = lRow + 1
                Next oResult
            End If
        End If
    
        ws.Select
    End With
    
    MsgBox "Process complete!", vbOKOnly + vbInformation, "Trade Simulation"
    
Exit_Here:

    Set ws = Nothing
    Set oResult = Nothing
    Set collFinalResults = Nothing
    Set oSimulation = Nothing
                
    Application.Calculation = iCalc
    Application.ScreenUpdating = blnScreenUpdating
    

End Sub


Function fncGetTrades() As Variant

'Purpose: return the input pnl trades as a one dimensional array
'         Now supports both simple PNL values and options trade strings

    Dim ws As Worksheet
    Dim rng As Range
    Dim arr As Variant
    Dim arrOptionsPNL As Variant
    Dim arrSimplePNL As Variant
    Dim arrCombined As Variant
    Dim lnglastRow As Long
    Dim lngfirstRow As Long
    Dim i As Long
    Dim j As Long
    Dim strCellValue As String
    Dim dblCurrentUnderlyingPrice As Double
    Dim intNumSimulations As Integer
    Dim dblVolatility As Double
    Dim lOptionsCount As Long
    Dim lSimpleCount As Long
    Dim lCombinedIndex As Long
    
    Set ws = ThisWorkbook.Worksheets("InputData")
    lnglastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lngfirstRow = 2
    
    ' Check if we have options trades or simple PNL
    ' First, separate options trades from simple PNL values
    lOptionsCount = 0
    lSimpleCount = 0
    
    ' Count options trades and simple PNL
    For i = lngfirstRow To lnglastRow
        strCellValue = Trim(CStr(ws.Cells(i, 1).value))
        If strCellValue <> "" Then
            If mdOptionsProcessor.IsOptionsTradeString(strCellValue) Then
                lOptionsCount = lOptionsCount + 1
            Else
                ' Try to parse as number
                If IsNumeric(strCellValue) Then
                    lSimpleCount = lSimpleCount + 1
                End If
            End If
        End If
    Next i
    
    ' Get underlying price and volatility from Control sheet (if available)
    On Error Resume Next
    dblCurrentUnderlyingPrice = ThisWorkbook.Sheets("Control").Range("UNDERLYING_PRICE").value
    If dblCurrentUnderlyingPrice = 0 Then
        ' Default to a reasonable value (e.g., QQQ around 600)
        dblCurrentUnderlyingPrice = 600
    End If
    
    intNumSimulations = ThisWorkbook.Sheets("Control").Range("OPTIONS_SIMULATIONS").value
    If intNumSimulations = 0 Then
        intNumSimulations = 1000 ' Default
    End If
    
    dblVolatility = ThisWorkbook.Sheets("Control").Range("UNDERLYING_VOLATILITY").value
    If dblVolatility = 0 Then
        dblVolatility = 0.2 ' Default 20% annual volatility
    End If
    On Error GoTo 0
    
    ' Process options trades if any
    If lOptionsCount > 0 Then
        Set rng = ws.Range("A" & lngfirstRow & ":A" & lnglastRow)
        arrOptionsPNL = mdOptionsProcessor.GetOptionsTradesAsPNL(ws, rng, dblCurrentUnderlyingPrice, intNumSimulations, dblVolatility)
    Else
        ReDim arrOptionsPNL(1 To 0)
    End If
    
    ' Get simple PNL values
    If lSimpleCount > 0 Then
        ReDim arrSimplePNL(1 To lSimpleCount)
        lCombinedIndex = 1
        For i = lngfirstRow To lnglastRow
            strCellValue = Trim(CStr(ws.Cells(i, 1).value))
            If strCellValue <> "" And Not mdOptionsProcessor.IsOptionsTradeString(strCellValue) Then
                If IsNumeric(strCellValue) Then
                    arrSimplePNL(lCombinedIndex) = CDbl(strCellValue)
                    lCombinedIndex = lCombinedIndex + 1
                End If
            End If
        Next i
    Else
        ReDim arrSimplePNL(1 To 0)
    End If
    
    ' Combine both arrays
    lCombinedIndex = UBound(arrOptionsPNL) - LBound(arrOptionsPNL) + 1 + UBound(arrSimplePNL) - LBound(arrSimplePNL) + 1
    If lCombinedIndex > 0 Then
        ReDim arrCombined(1 To lCombinedIndex)
        lCombinedIndex = 1
        
        ' Add options PNL
        For i = LBound(arrOptionsPNL) To UBound(arrOptionsPNL)
            arrCombined(lCombinedIndex) = arrOptionsPNL(i)
            lCombinedIndex = lCombinedIndex + 1
        Next i
        
        ' Add simple PNL
        For i = LBound(arrSimplePNL) To UBound(arrSimplePNL)
            arrCombined(lCombinedIndex) = arrSimplePNL(i)
            lCombinedIndex = lCombinedIndex + 1
        Next i
        
        fncGetTrades = arrCombined
    Else
        ' Fallback to original method if no valid trades found
        Set rng = ws.Range("A" & lngfirstRow & ":A" & lnglastRow)
        fncGetTrades = Application.Transpose(rng.value)
    End If
    
    Set ws = Nothing
    Set rng = Nothing

End Function


Public Sub ClearUI()

'Purpose: Reset this tool and any input ranges

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Control")
    ws.Range("OUTPUT").ClearContents
    
    Set ws = Nothing

End Sub
