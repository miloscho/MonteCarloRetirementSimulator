Attribute VB_Name = "Module1"
Option Explicit

Sub RunPortfolioMonteCarloSimulation( _
        Optional initialPortfolio As Double = 1000000, _
        Optional initialWithdrawal As Double = 40000, _
        Optional retirementYears As Long = 30, _
        Optional numSimulations As Long = 10000, _
        Optional stockAllocPct As Double = 60, _
        Optional annualFeesPct As Double = 0.8, _
        Optional blockSize As Long = 5, _
        Optional useBlocks As Boolean = True, _
        Optional ssAmount As Double = 0, _
        Optional ssStartYear As Long = 0)

    Dim wsData As Worksheet, wsResults As Worksheet
    Dim wsPath As Worksheet, wsHist As Worksheet
   
    Set wsData = ThisWorkbook.Sheets("HistoricalData")
   
    ' Create or clear Results sheet
    On Error Resume Next
    Set wsResults = ThisWorkbook.Sheets("Results")
    On Error GoTo 0
    If wsResults Is Nothing Then
        Set wsResults = ThisWorkbook.Sheets.Add
        wsResults.Name = "Results"
    Else
        wsResults.Cells.Clear
    End If
   
    ' === CONVERT PERCENTAGES ===
    Dim stockAlloc As Double, annualFees As Double
    stockAlloc = stockAllocPct / 100
    annualFees = annualFeesPct / 100
   
    If blockSize < 1 Then blockSize = 1
    If ssStartYear < 1 Then ssStartYear = 999
   
    ' === LOAD HISTORICAL DATA ===
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    Dim numYears As Long
    numYears = lastRow - 1
  
    Dim histSP() As Double, histBonds() As Double, histInfl() As Double
    ReDim histSP(1 To numYears), histBonds(1 To numYears), histInfl(1 To numYears)
  
    Dim i As Long
    For i = 1 To numYears
        histSP(i) = wsData.Cells(i + 1, 2).Value
        histBonds(i) = wsData.Cells(i + 1, 4).Value
        histInfl(i) = wsData.Cells(i + 1, 3).Value
    Next i
  
    ' === RUN MONTE CARLO ===
    Randomize
    
    'Initialize tracking variables
    Dim successes As Long, successFinalsSum As Double
    Dim finalPortfolios() As Double
    ReDim finalPortfolios(1 To numSimulations)
  
    Dim sim As Long, yr As Long
    Dim randStart As Long, blockStart As Long, currentBlockPos As Long
    Dim portfolio As Double, withdrawal As Double
    Dim blendedReturn As Double, thisInfl As Double
    Dim isSuccess As Boolean
    Dim totalWithdrawal As Double
  
    
    ' Store the path of the first simulation for charting
    Dim pathPortfolio() As Double
    ReDim pathPortfolio(0 To retirementYears)
    pathPortfolio(0) = initialPortfolio
  
    successes = 0
    successFinalsSum = 0
  
    Application.ScreenUpdating = False
    Application.StatusBar = "Running " & numSimulations & " simulations... 0% complete"
  
    ' ====================== MAIN MONTE CARLO LOOP ======================
    For sim = 1 To numSimulations
        portfolio = initialPortfolio
        withdrawal = initialWithdrawal
        isSuccess = True
        currentBlockPos = 1
      
        ' ====================== YEAR-BY-YEAR SIMULATION ======================
        
        For yr = 1 To retirementYears
            
            ' Select historical year (with optional block bootstrapping)
            If useBlocks And blockSize > 1 Then
                If currentBlockPos = 1 Then
                    blockStart = Int(Rnd * (numYears - blockSize + 1)) + 1
                End If
                randStart = blockStart + currentBlockPos - 1
                currentBlockPos = currentBlockPos + 1
                If currentBlockPos > blockSize Then currentBlockPos = 1
            Else
                randStart = Int(Rnd * numYears) + 1
            End If
          
            ' Calculate blended portfolio return and inflation
            blendedReturn = stockAlloc * histSP(randStart) + (1 - stockAlloc) * histBonds(randStart) - annualFees
            thisInfl = histInfl(randStart)
          
            ' Calculate withdrawal (adjusted for Social Security if applicable)
            totalWithdrawal = withdrawal
            If yr >= ssStartYear And ssAmount > 0 Then
                Dim ssThisYear As Double
                ssThisYear = ssAmount * (1 + thisInfl) ^ (yr - ssStartYear)
                totalWithdrawal = withdrawal - ssThisYear
                If totalWithdrawal < 0 Then totalWithdrawal = 0
            End If
          
            ' Check for portfolio failure
            If portfolio < totalWithdrawal Then
                isSuccess = False
                Exit For
            End If
          
            ' Update portfolio balance
            portfolio = portfolio - totalWithdrawal
            portfolio = portfolio * (1 + blendedReturn)
            
            ' Inflate next year's withdrawal (for fixed withdrawal method)
            withdrawal = withdrawal * (1 + thisInfl)
          
            ' Save path for first simulation only (for charting)
            If sim = 1 Then pathPortfolio(yr) = portfolio
        Next yr
      
        ' Record simulation result
        If isSuccess Then
            successes = successes + 1
            successFinalsSum = successFinalsSum + portfolio
            finalPortfolios(sim) = portfolio
        Else
            finalPortfolios(sim) = 0
        End If
      
        ' Progress feedback every 500 simulations
        If sim Mod 500 = 0 Then
            Application.StatusBar = "Running simulations... " & Format(sim / numSimulations, "0%") & " complete"
            DoEvents
        End If
    Next sim
  
    Application.StatusBar = False
    Application.ScreenUpdating = True
  
   
    ' === WRITE RESULTS SHEET ===
    Dim successRate As Double
    successRate = successes / numSimulations
   
    Dim avgFinalAll As Double, avgFinalSurvivors As Double
    avgFinalAll = Application.WorksheetFunction.Average(finalPortfolios)
    If successes > 0 Then avgFinalSurvivors = successFinalsSum / successes
   
    With wsResults
        .Range("A1").Value = "Monte Carlo Simulator with Social Security"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
       
        .Range("A3").Value = "Success Rate"
        .Range("B3").Value = successRate
        .Range("B3").NumberFormat = "0.00%"
       
        .Range("A4").Value = "Successful Simulations"
        .Range("B4").Value = successes & " / " & numSimulations
       
        .Range("A5").Value = "Avg Ending Portfolio (all)"
        .Range("B5").Value = avgFinalAll
        .Range("B5").NumberFormat = "$#,##0"
       
        .Range("A6").Value = "Avg Ending (survivors)"
        .Range("B6").Value = avgFinalSurvivors
        .Range("B6").NumberFormat = "$#,##0"
       
        .Range("A8").Value = "Ending Portfolio Percentiles"
        Dim pcts As Variant
        pcts = Array(0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95)
        Dim lbls As Variant
        lbls = Array("5th", "10th", "25th", "50th (Median)", "75th", "90th", "95th")
       
        Dim j As Long
        For j = 0 To 6
            .Cells(9 + j, 1).Value = lbls(j)
            .Cells(9 + j, 2).Value = Application.WorksheetFunction.Percentile(finalPortfolios, pcts(j))
            .Cells(9 + j, 2).NumberFormat = "$#,##0"
        Next j
        .Columns("A:B").AutoFit
    End With
   
    ' === THIRD CHART: SUMMARY BAR CHART (Success Rate Removed) ===
    Dim chSummary As ChartObject
    Set chSummary = wsResults.ChartObjects.Add(Left:=wsResults.Range("D3").Left, _
                                               Top:=wsResults.Range("D3").Top, _
                                               Width:=650, Height:=380)
    
    With chSummary.Chart
        .ChartType = xlColumnClustered
        
        ' Prepare data (without Success Rate)
        wsResults.Range("D8").Value = "Metric"
        wsResults.Range("E8").Value = "Value"
        
        wsResults.Range("D9").Value = "Median Ending Balance"
        wsResults.Range("E9").Value = Application.WorksheetFunction.Percentile(finalPortfolios, 0.5)
        
        wsResults.Range("D10").Value = "5th Percentile (Worst Case)"
        wsResults.Range("E10").Value = Application.WorksheetFunction.Percentile(finalPortfolios, 0.05)
        
        wsResults.Range("D11").Value = "Average Ending Balance"
        wsResults.Range("E11").Value = avgFinalAll
        
        .SetSourceData Source:=wsResults.Range("D8:E11")
        .HasTitle = True
        .ChartTitle.Text = "Key Simulation Outcomes"
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Portfolio Value ($)"
        .Legend.Delete
    End With
   
    ' === FIRST SIMULATION PATH ===
    On Error Resume Next
    Set wsPath = ThisWorkbook.Sheets("First Simulation Path")
    On Error GoTo 0
    If wsPath Is Nothing Then
        Set wsPath = ThisWorkbook.Sheets.Add(After:=wsResults)
        wsPath.Name = "First Simulation Path"
    Else
        wsPath.Cells.Clear
        Dim chObj As ChartObject
        For Each chObj In wsPath.ChartObjects
            chObj.Delete
        Next chObj
    End If
   
    wsPath.Range("A1").Value = "Year"
    wsPath.Range("B1").Value = "Portfolio Balance"
    For i = 0 To retirementYears
        wsPath.Cells(i + 2, 1).Value = i
        wsPath.Cells(i + 2, 2).Value = pathPortfolio(i)
    Next i
    wsPath.Columns("B").NumberFormat = "$#,##0"
   
    Dim chPath As ChartObject
    Set chPath = wsPath.ChartObjects.Add(Left:=wsPath.Range("D3").Left, Top:=wsPath.Range("D3").Top, Width:=720, Height:=420)
    With chPath.Chart
        .ChartType = xlLine
        .SetSourceData Source:=wsPath.Range("A1:B" & retirementYears + 2)
        .HasTitle = True
        .ChartTitle.Text = "Portfolio Balance Over Time - First Simulation"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Year"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Portfolio Value ($)"
        .Legend.Delete
    End With
   
    ' === OUTCOME HISTOGRAM ===
    On Error Resume Next
    Set wsHist = ThisWorkbook.Sheets("Outcome Histogram")
    On Error GoTo 0
    If wsHist Is Nothing Then
        Set wsHist = ThisWorkbook.Sheets.Add(After:=wsPath)
        wsHist.Name = "Outcome Histogram"
    Else
        wsHist.Cells.Clear
        For Each chObj In wsHist.ChartObjects
            chObj.Delete
        Next chObj
    End If
   
    Dim binWidth As Double: binWidth = 1000000
    Dim maxBin As Long: maxBin = 10
    Dim b As Long, r As Long
   
    wsHist.Range("A1").Value = "Range"
    wsHist.Range("B1").Value = "Number of Simulations"
   
    For b = 1 To maxBin
        wsHist.Cells(b + 1, 1).Value = (b - 1) * binWidth & " - " & b * binWidth
        wsHist.Cells(b + 1, 1).NumberFormat = "$#,##0"
    Next b
    wsHist.Cells(maxBin + 2, 1).Value = "10M+"
   
    Dim freq() As Long
    ReDim freq(1 To maxBin + 1)
   
    For r = 1 To numSimulations
        Dim binIndex As Long
        binIndex = Application.WorksheetFunction.Min(Int(finalPortfolios(r) / binWidth) + 1, maxBin + 1)
        freq(binIndex) = freq(binIndex) + 1
    Next r
   
    For b = 1 To maxBin + 1
        wsHist.Cells(b + 1, 2).Value = freq(b)
    Next b
   
    Dim chHist As ChartObject
    Set chHist = wsHist.ChartObjects.Add(Left:=wsHist.Range("D3").Left, Top:=wsHist.Range("D3").Top, Width:=720, Height:=420)
    With chHist.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsHist.Range("A1:B" & maxBin + 2)
        .HasTitle = True
        .ChartTitle.Text = "Distribution of Ending Portfolio Values"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Ending Portfolio Value Range"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Number of Simulations"
        .Legend.Delete
        On Error Resume Next
        .ChartGroups(1).GapWidth = 0
        On Error GoTo 0
    End With
   
    wsHist.Columns("A:B").AutoFit
   
    ' Show Results sheet
    wsResults.Activate
    wsResults.Range("A1").Select
   
    'MsgBox "Simulation Complete!" & vbCrLf & _
         '  "Success Rate: " & Format(successRate, "0.00%") & vbCrLf & vbCrLf & _
           ' "You now have 3 charts:" & vbCrLf & _
           ' "1. Summary Bar Chart (on Results sheet)" & vbCrLf & _
          '  "2. First Simulation Path" & vbCrLf & _
           ' "3. Outcome Histogram", vbInformation
        
        
        ' Copy everything to Inputs sheet
    Call CopyChartsToInputsSheet( _
        initialPortfolio, initialWithdrawal, retirementYears, numSimulations, _
        stockAllocPct, annualFeesPct, blockSize, useBlocks, ssAmount, ssStartYear)
    
   ' MsgBox "Simulation Complete!" & vbCrLf & _
        '   "Success Rate: " & Format(successRate, "0.00%"), vbInformation
    MsgBox "Results have been updated."
           
    
    
End Sub


Sub AddSimulationButton()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Inputs")
    
    ' Delete old button if it already exists (to avoid duplicates)
    On Error Resume Next
    ws.Shapes("RunMonteCarloButton").Delete
    On Error GoTo 0
    
    ' Create new button
    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
                Left:=ws.Range("D1").Left, _
                Top:=ws.Range("D1").Top, _
                Width:=220, _
                Height:=45)
    
    With btn
        .Name = "RunMonteCarloButton"
        .TextFrame.Characters.Text = "RUN MONTE CARLO SIMULATION"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(0, 102, 204)      ' Nice blue color
        .Line.ForeColor.RGB = RGB(0, 51, 102)
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
    End With
    
    ' Assign the macro to the button
    btn.OnAction = "RunPortfolioMonteCarloSimulation"
    
    MsgBox "? Button successfully added to the Inputs sheet!" & vbCrLf & vbCrLf & _
           "You can now click the big blue button to run the simulation anytime.", vbInformation

End Sub

Sub LaunchRetirementSimulator()
    frmRetirementSimulator.Show
End Sub

Private Sub CopyChartsToInputsSheet( _
        initialPortfolio As Double, initialWithdrawal As Double, _
        retirementYears As Long, numSimulations As Long, _
        stockAllocPct As Double, annualFeesPct As Double, _
        blockSize As Long, useBlocks As Boolean, _
        ssAmount As Double, ssStartYear As Long)

    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets("Inputs")
    
    ' Clear previous charts
    Dim ch As ChartObject
    For Each ch In wsInputs.ChartObjects
        ch.Delete
    Next ch
    
    ' Clear previous text areas
    wsInputs.Range("D1:G200").Clear
    wsInputs.Range("H1:Z200").Clear
    
    ' ====================== LATEST SIMULATION VALUES (Top Right) ======================
    wsInputs.Range("D2").Value = "LATEST SIMULATION VALUES"
    wsInputs.Range("D2").Font.Bold = True
    wsInputs.Range("D2").Font.Size = 14
    
    Dim row As Long: row = 4
    
    With wsInputs
        .Cells(row, 4).Value = "Initial Portfolio":     .Cells(row, 5).Value = initialPortfolio:     .Cells(row, 5).NumberFormat = "$#,##0": row = row + 1
        .Cells(row, 4).Value = "Initial Withdrawal":    .Cells(row, 5).Value = initialWithdrawal:    .Cells(row, 5).NumberFormat = "$#,##0": row = row + 1
        .Cells(row, 4).Value = "Retirement Years":      .Cells(row, 5).Value = retirementYears:      row = row + 1
        .Cells(row, 4).Value = "Number of Simulations": .Cells(row, 5).Value = numSimulations:       row = row + 1
        .Cells(row, 4).Value = "Stock Allocation":      .Cells(row, 5).Value = stockAllocPct / 100:  .Cells(row, 5).NumberFormat = "0.0%": row = row + 1
        .Cells(row, 4).Value = "Annual Fees":           .Cells(row, 5).Value = annualFeesPct / 100:  .Cells(row, 5).NumberFormat = "0.00%": row = row + 1
        .Cells(row, 4).Value = "Block Size":            .Cells(row, 5).Value = blockSize:            row = row + 1
        .Cells(row, 4).Value = "Use Block Bootstrapping": .Cells(row, 5).Value = IIf(useBlocks, "Yes", "No"): row = row + 1
        .Cells(row, 4).Value = "Social Security Amount": .Cells(row, 5).Value = ssAmount:            .Cells(row, 5).NumberFormat = "$#,##0": row = row + 1
        .Cells(row, 4).Value = "SS Start Year":         .Cells(row, 5).Value = ssStartYear:          row = row + 1
    End With
    
    ' ====================== SIMULATION RESULTS (Far Right) ======================
    wsInputs.Range("H2").Value = "SIMULATION RESULTS"
    wsInputs.Range("H2").Font.Bold = True
    wsInputs.Range("H2").Font.Size = 14
    
    wsInputs.Range("H4").Value = "Success Rate:"
    wsInputs.Range("I4").Value = ThisWorkbook.Sheets("Results").Range("B3").Value
    wsInputs.Range("I4").NumberFormat = "0.00%"
    
    wsInputs.Range("H5").Value = "Successful Simulations:"
    wsInputs.Range("I5").Value = ThisWorkbook.Sheets("Results").Range("B4").Value
    
    wsInputs.Range("H6").Value = "Avg Ending Portfolio (All):"
    wsInputs.Range("I6").Value = ThisWorkbook.Sheets("Results").Range("B5").Value
    wsInputs.Range("I6").NumberFormat = "$#,##0"
    
    wsInputs.Range("H7").Value = "Avg Ending (Survivors):"
    wsInputs.Range("I7").Value = ThisWorkbook.Sheets("Results").Range("B6").Value
    wsInputs.Range("I7").NumberFormat = "$#,##0"
    
    wsInputs.Range("H9").Value = "ENDING PORTFOLIO PERCENTILES"
    wsInputs.Range("H9").Font.Bold = True
    
    Dim i As Long
    For i = 0 To 6
        wsInputs.Cells(10 + i, 8).Value = ThisWorkbook.Sheets("Results").Cells(9 + i, 1).Value
        wsInputs.Cells(10 + i, 9).Value = ThisWorkbook.Sheets("Results").Cells(9 + i, 2).Value
        wsInputs.Cells(10 + i, 9).NumberFormat = "$#,##0"
    Next i
    
    ' ====================== CHARTS - VERTICAL STACK ON LEFT ======================
    Dim leftPos As Double:    leftPos = 20          ' Far left
    Dim topPos As Double:     topPos = 420          ' First chart starts ~ row 30
    Dim chartWidth As Double: chartWidth = 720
    Dim chartHeight As Double: chartHeight = 340
    Dim verticalGap As Double: verticalGap = 40
    
    ' Chart 1: Summary Bar Chart (starts at ~row 30)
    Dim ch1 As ChartObject
    Set ch1 = wsInputs.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=chartWidth, Height:=chartHeight)
    With ch1.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ThisWorkbook.Sheets("Results").Range("D8:E11")
        .HasTitle = True
        .ChartTitle.Text = "Key Simulation Outcomes"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Portfolio Value ($)"
        .Legend.Delete
    End With
    
    ' Chart 2: Portfolio Path
    Dim ch2 As ChartObject
    Set ch2 = wsInputs.ChartObjects.Add(Left:=leftPos, Top:=topPos + chartHeight + verticalGap, _
                                        Width:=chartWidth, Height:=chartHeight)
    With ch2.Chart
        .ChartType = xlLine
        .SetSourceData Source:=ThisWorkbook.Sheets("First Simulation Path").Range("A1:B" & _
                        ThisWorkbook.Sheets("First Simulation Path").Cells(Rows.Count, "A").End(xlUp).row)
        .HasTitle = True
        .ChartTitle.Text = "Portfolio Path - First Simulation"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Year"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Portfolio Value ($)"
        .Legend.Delete
    End With
    
    ' Chart 3: Outcome Histogram
    Dim ch3 As ChartObject
    Set ch3 = wsInputs.ChartObjects.Add(Left:=leftPos, Top:=topPos + 2 * (chartHeight + verticalGap), _
                                        Width:=chartWidth, Height:=chartHeight)
    With ch3.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ThisWorkbook.Sheets("Outcome Histogram").Range("A1:B" & _
                        ThisWorkbook.Sheets("Outcome Histogram").Cells(Rows.Count, "A").End(xlUp).row)
        .HasTitle = True
        .ChartTitle.Text = "Distribution of Ending Portfolios"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Ending Value Range"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Number of Simulations"
        .Legend.Delete
    End With
    
    wsInputs.Activate
    wsInputs.Range("A1").Select
    
End Sub
