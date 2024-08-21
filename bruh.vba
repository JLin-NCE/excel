Option Explicit

Private Type InputDataType
    streetName As String
    startLocation As String
    endLocation As String
    cutLength As Double
    cutWidth As Double
    distanceFromPrevSection As Double
    anticipatedCutYear As Integer
    inflationRate As Double
End Type

Public Sub GatherAssociatedRows()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim wsPCI As Worksheet
    Dim inputData As InputDataType
    Dim startRow As Long, endRow As Long
    
    Set wsInput = ThisWorkbook.Sheets("Cut Impact Fee Calculator")
    Set wsPCI = ThisWorkbook.Sheets("Covina PCI Report")
    Set wsOutput = CreateOrClearOutputSheet
    
    inputData = GetInputData(wsInput)
    
    If Not FindStartAndEndRows(wsPCI, inputData, startRow, endRow) Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ProcessAndOutputData wsOutput, wsPCI, wsInput, inputData, startRow, endRow
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Data extraction complete from row " & startRow & " to row " & endRow, vbInformation, "Success"
End Sub

Private Function CreateOrClearOutputSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Fee Calculator Output")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Cut Impact Fee Calculator"))
        ws.Name = "Fee Calculator Output"
    Else
        ws.Cells.Clear
    End If
    
    With ws
        .Cells(1).Resize(1, 16).Value = Array("Street Name", "From", "To", "Section Start", "Section End", "Length", "Width", "Area", "PCI", "Functional Class", "Cut Type", "Cut Area", "Small Cut Fee", "Large Cut Fee", "Fee Calculation", "Cut Cost")
    End With
    
    Set CreateOrClearOutputSheet = ws
End Function

Private Function GetInputData(ws As Worksheet) As InputDataType
    Dim data As InputDataType
    
    With ws
        data.streetName = .Cells(3, 3).Value
        data.startLocation = .Cells(4, 3).Value
        data.endLocation = .Cells(5, 3).Value
        data.cutLength = Round(CDbl(.Cells(6, 3).Value), 2)
        data.cutWidth = Round(CDbl(.Cells(7, 3).Value), 2)
        data.distanceFromPrevSection = Round(CDbl(.Cells(8, 3).Value), 2)
        data.anticipatedCutYear = CInt(.Cells(9, 3).Value)
        data.inflationRate = Round(CDbl(.Cells(10, 3).Value), 2)
    End With
    
    GetInputData = data
End Function

Private Function FindStartAndEndRows(ws As Worksheet, inputData As InputDataType, ByRef startRow As Long, ByRef endRow As Long) As Boolean
    Dim dataArray As Variant
    Dim i As Long, lastRow As Long
    Dim totalLength As Double
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    dataArray = ws.Range("C1:J" & lastRow).Value
    
    For i = 2 To UBound(dataArray)
        If dataArray(i, 1) = inputData.streetName And dataArray(i, 2) = inputData.startLocation Then
            startRow = i
            Exit For
        End If
    Next i
    
    If startRow = 0 Then
        MsgBox "Beginning location not found for Street Name: " & inputData.streetName & ", From: " & inputData.startLocation, vbExclamation, "Error"
        FindStartAndEndRows = False
        Exit Function
    End If
    
    totalLength = 0
    For i = startRow To UBound(dataArray)
        If dataArray(i, 1) = inputData.streetName Then
            totalLength = totalLength + dataArray(i, 8) ' Column J is Length
            
            If dataArray(i, 3) = "END" Or dataArray(i, 3) = inputData.endLocation Then
                endRow = i
                FindStartAndEndRows = True
                Exit Function
            End If
            
            If totalLength >= inputData.distanceFromPrevSection + inputData.cutLength Then
                endRow = i
                FindStartAndEndRows = True
                Exit Function
            End If
        Else
            MsgBox "Error: End location not found before street name changed." & vbNewLine & _
                   "Street Name: " & inputData.streetName & vbNewLine & _
                   "Specified End Location: " & inputData.endLocation, vbExclamation, "Error"
            FindStartAndEndRows = False
            Exit Function
        End If
    Next i
    
    MsgBox "Ending location not found for Street Name: " & inputData.streetName & ", To: " & inputData.endLocation, vbExclamation, "Error"
    FindStartAndEndRows = False
End Function

Private Sub FormatOutputSheet(ws As Worksheet, lastRow As Long)
    If lastRow < 2 Then Exit Sub  ' Exit if there's no data
    
    With ws
        ' Clear any existing formatting
        .Cells.ClearFormats
        
        ' Determine the last column
        Dim lastCol As Long
        lastCol = .Cells(2, .Columns.Count).End(xlToLeft).Column
        
        ' Add a title to the sheet
        .Cells(1, 1).EntireRow.Insert
        With .Cells(1, 1)
            .Value = "Road Cut Fee Calculation Results"
            .Font.Size = 14
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        If lastCol > 1 Then
            .Range(.Cells(1, 1), .Cells(1, lastCol)).Merge
        End If
        
        ' Format headers
        With .Range(.Cells(2, 1), .Cells(2, lastCol))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(155, 194, 230)  ' Light blue
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
        
        ' Format all data rows excluding the "Total Cut Cost" row
        With .Range(.Cells(3, 1), .Cells(lastRow - 1, lastCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Bold = False  ' Ensure no bold font for data rows
        End With
        
        ' Add alternating row colors and ensure consistent shading
        Dim i As Long
        For i = 3 To lastRow - 1
            If i Mod 2 = 0 Then
                .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.Color = RGB(217, 225, 242)  ' Very light blue
            Else
                .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.Color = RGB(255, 255, 255)  ' White for odd rows
            End If
        Next i
        
        ' Apply consistent formatting to all relevant columns
        .Range(.Cells(3, 4), .Cells(lastRow - 1, 8)).NumberFormat = "#,##0.00"  ' Section Start, Section End, Length, Width, Area
        .Range(.Cells(3, 9), .Cells(lastRow - 1, 9)).NumberFormat = "0.0"  ' PCI
        .Range(.Cells(3, 12), .Cells(lastRow - 1, 12)).NumberFormat = "#,##0.00"  ' Cut Area
        .Range(.Cells(3, 13), .Cells(lastRow - 1, 14)).NumberFormat = "$#,##0.00"  ' Small Cut Fee, Large Cut Fee
        .Range(.Cells(3, 15), .Cells(lastRow - 1, 15)).NumberFormat = "$0 * #,##0.00"  ' Fee Calculation (with $ sign)
        .Range(.Cells(3, 16), .Cells(lastRow - 1, 18)).NumberFormat = "$#,##0.00"  ' Cut Cost and new adjusted columns
        
        ' Bold and highlight the "Total Cut Cost" row
        With .Range(.Cells(lastRow, 1), .Cells(lastRow, lastCol))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter  ' Ensure consistent alignment in the total row
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(189, 215, 238)  ' Medium blue background for clarity
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' Ensure row shading is consistent even for the last row
        If lastRow Mod 2 = 0 Then
            .Range(.Cells(lastRow, 1), .Cells(lastRow, lastCol)).Interior.Color = RGB(217, 225, 242)  ' Light blue for last row
        Else
            .Range(.Cells(lastRow, 1), .Cells(lastRow, lastCol)).Interior.Color = RGB(255, 255, 255)  ' White for last row
        End If
        
        ' Add a blank row after the last data row
        .Cells(lastRow + 1, 1).EntireRow.Insert
        
        ' Adjust column widths consistently for all columns
        .UsedRange.Columns.AutoFit
    End With
End Sub





Private Sub ProcessAndOutputData(wsOutput As Worksheet, wsPCI As Worksheet, wsInput As Worksheet, inputData As InputDataType, startRow As Long, endRow As Long)
    Dim dataArray As Variant
    Dim outputArray() As Variant
    Dim i As Long, outputRow As Long
    Dim remainingCutLength As Double, currentCutStart As Double, totalCutCost As Double
    Dim sectionLength As Double, sectionWidth As Double, sectionArea As Double
    Dim sectionStart As Double, sectionEnd As Double
    Dim functionalClass As String, functionalClassFull As String, pci As Double
    Dim smallCutFee As Double, largeCutFee As Double
    Dim cutArea As Double, cutCost As Double, cutType As String, feeCalculation As String
    Dim inflationFactor As Double
    
    ' Calculate the inflation factor based on the anticipated cut year and inflation rate
    If inputData.anticipatedCutYear > 2024 Then
        inflationFactor = (1 + inputData.inflationRate / 100) ^ (inputData.anticipatedCutYear - 2024)
    Else
        inflationFactor = 1 ' No inflation adjustment if the year is 2024 or earlier
    End If
    
    dataArray = wsPCI.Range("C" & startRow & ":N" & endRow).Value
    ReDim outputArray(1 To endRow - startRow + 1, 1 To 18)  ' Increased to 18 to include Adjusted Small and Large Cut Fee
    
    outputRow = 1
    remainingCutLength = inputData.cutLength
    currentCutStart = inputData.distanceFromPrevSection
    totalCutCost = 0
    
    For i = 1 To UBound(dataArray)
        sectionLength = Round(dataArray(i, 8), 2)  ' Column J
        sectionWidth = Round(dataArray(i, 9), 2)   ' Column K
        sectionArea = Round(dataArray(i, 10), 2)   ' Column L
        
        If i = 1 Then
            sectionStart = Round(inputData.distanceFromPrevSection, 2)
            sectionEnd = Round(sectionLength, 2)
            sectionLength = Round(sectionEnd - sectionStart, 2)
        Else
            sectionStart = 0
            sectionEnd = Round(sectionLength, 2)
        End If
        
        If remainingCutLength <= sectionLength Then
            sectionEnd = Round(sectionStart + remainingCutLength, 2)
            sectionLength = Round(remainingCutLength, 2)
        End If
        
        functionalClass = dataArray(i, 6)  ' Column H
        pci = Round(dataArray(i, 12), 2)   ' Column N
        cutArea = Round(sectionLength * inputData.cutWidth, 2)
        
        DetermineFees functionalClass, pci, functionalClassFull, smallCutFee, largeCutFee
        
        ' Adjust the cut fees using the inflation factor
        smallCutFee = Round(smallCutFee * inflationFactor, 2)
        largeCutFee = Round(largeCutFee * inflationFactor, 2)
        
        DetermineCutTypeAndCost cutArea, sectionArea, smallCutFee, largeCutFee, cutType, cutCost, feeCalculation
        
        outputArray(outputRow, 1) = dataArray(i, 1)  ' Street Name
        outputArray(outputRow, 2) = dataArray(i, 2)  ' From
        outputArray(outputRow, 3) = dataArray(i, 3)  ' To
        outputArray(outputRow, 4) = Round(sectionStart, 2)
        outputArray(outputRow, 5) = Round(sectionEnd, 2)
        outputArray(outputRow, 6) = Round(sectionLength, 2)
        outputArray(outputRow, 7) = Round(sectionWidth, 2)
        outputArray(outputRow, 8) = Round(sectionLength * sectionWidth, 2)
        outputArray(outputRow, 9) = Round(pci, 2)
        outputArray(outputRow, 10) = functionalClassFull
        outputArray(outputRow, 11) = cutType
        outputArray(outputRow, 12) = Round(cutArea, 2)
        outputArray(outputRow, 13) = "$" & Format(smallCutFee, "#,##0.00")  ' Adjusted Small Cut Fee with $
        outputArray(outputRow, 14) = "$" & Format(largeCutFee, "#,##0.00")  ' Adjusted Large Cut Fee with $
        outputArray(outputRow, 15) = feeCalculation
        outputArray(outputRow, 16) = "$" & Format(cutCost, "#,##0.00")
        outputArray(outputRow, 17) = "$" & Format(smallCutFee, "#,##0.00")  ' New column for Adjusted Small Cut Fee
        outputArray(outputRow, 18) = "$" & Format(largeCutFee, "#,##0.00")  ' New column for Adjusted Large Cut Fee
        
        remainingCutLength = Round(remainingCutLength - sectionLength, 2)
        currentCutStart = Round(sectionEnd, 2)
        totalCutCost = Round(totalCutCost + cutCost, 2)
        
        If remainingCutLength <= 0 Then Exit For
        
        outputRow = outputRow + 1
    Next i
    
    ' Handling case where cut length exceeds the available sections
    If remainingCutLength > 0 Then
        MsgBox "Warning: The cut length exceeds the available sections. Remaining cut length: " & remainingCutLength, vbExclamation, "Warning"
    End If
    
    wsOutput.Cells(2, 1).Resize(outputRow, 18).Value = outputArray
    
    outputRow = outputRow + 2
    wsOutput.Cells(outputRow, 1).Value = "Total Cut Cost"
    wsOutput.Cells(outputRow, 16).Value = "$" & Format(totalCutCost, "#,##0.00")
    
    ' Output the total cut cost to C11 on the "Cut Impact Fee Calculator" sheet
    wsInput.Cells(11, 3).Value = totalCutCost
    
    FormatOutputSheet wsOutput, outputRow
End Sub



Private Sub DetermineFees(functionalClass As String, pci As Double, ByRef functionalClassFull As String, ByRef smallCutFee As Double, ByRef largeCutFee As Double)
    Select Case functionalClass
        Case "A"
            functionalClassFull = "Arterials"
            If pci >= 70 Then
                smallCutFee = 1
                largeCutFee = 4.5
            Else
                smallCutFee = 0.5
                largeCutFee = 0.5
            End If
        Case "C"
            functionalClassFull = "Collectors"
            If pci >= 70 Then
                smallCutFee = 1
                largeCutFee = 4.5
            Else
                smallCutFee = 0.5
                largeCutFee = 0.5
            End If
        Case "E"
            functionalClassFull = "Residentials"
            If pci >= 50 Then
                smallCutFee = 1.5
                largeCutFee = 4
            Else
                smallCutFee = 0.25
                largeCutFee = 0.5
            End If
        Case Else
            functionalClassFull = "Unknown"
            smallCutFee = 0
            largeCutFee = 0
    End Select
End Sub

Private Sub DetermineCutTypeAndCost(cutArea As Double, sectionArea As Double, smallCutFee As Double, largeCutFee As Double, _
                                    ByRef cutType As String, ByRef cutCost As Double, ByRef feeCalculation As String)
    If cutArea < Round(0.1 * sectionArea, 2) Then
        cutType = "Small Cut"
        cutCost = Round(cutArea * smallCutFee, 2)
        feeCalculation = cutArea & " * " & smallCutFee
    Else
        cutType = "Large Cut"
        cutCost = Round(cutArea * largeCutFee, 2)
        feeCalculation = cutArea & " * " & largeCutFee
    End If
End Sub


