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
    
    Dim wsSheet3 As Worksheet
    Dim wsSheet3Output As Worksheet
    Dim wsPCI As Worksheet
    Dim inputData As InputDataType
    Dim startRow As Long, endRow As Long
    
    Set wsSheet3 = ThisWorkbook.Sheets("Sheet3")
    Set wsPCI = ThisWorkbook.Sheets("Covina PCI Report")
    Set wsSheet3Output = CreateOrClearOutputSheet
    
    inputData = GetInputData(wsSheet3)
    
    If Not FindStartAndEndRows(wsPCI, inputData, startRow, endRow) Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ProcessAndOutputData wsSheet3Output, wsPCI, inputData, startRow, endRow
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Data extraction complete from row " & startRow & " to row " & endRow, vbInformation, "Success"
End Sub

Private Function CreateOrClearOutputSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Sheet3 Output")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Sheet3"))
        ws.Name = "Sheet3 Output"
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

Private Sub ProcessAndOutputData(wsOutput As Worksheet, wsPCI As Worksheet, inputData As InputDataType, startRow As Long, endRow As Long)
    Dim dataArray As Variant
    Dim outputArray() As Variant
    Dim i As Long, outputRow As Long
    Dim remainingCutLength As Double, currentCutStart As Double, totalCutCost As Double
    Dim sectionLength As Double, sectionWidth As Double, sectionArea As Double
    Dim sectionStart As Double, sectionEnd As Double
    Dim functionalClass As String, functionalClassFull As String, pci As Double
    Dim smallCutFee As Double, largeCutFee As Double
    Dim cutArea As Double, cutCost As Double, cutType As String, feeCalculation As String
    
    dataArray = wsPCI.Range("C" & startRow & ":N" & endRow).Value
    ReDim outputArray(1 To endRow - startRow + 1, 1 To 16)
    
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
        outputArray(outputRow, 13) = Round(smallCutFee, 2)
        outputArray(outputRow, 14) = Round(largeCutFee, 2)
        outputArray(outputRow, 15) = feeCalculation
        outputArray(outputRow, 16) = Round(cutCost, 2)
        
        remainingCutLength = Round(remainingCutLength - sectionLength, 2)
        currentCutStart = Round(sectionEnd, 2)
        totalCutCost = Round(totalCutCost + cutCost, 2)
        
        If remainingCutLength <= 0 Or dataArray(i, 3) = "END" Then Exit For
        
        outputRow = outputRow + 1
    Next i
    
    wsOutput.Cells(2, 1).Resize(outputRow, 16).Value = outputArray
    
    outputRow = outputRow + 2
    wsOutput.Cells(outputRow, 1).Value = "Total Cut Cost"
    wsOutput.Cells(outputRow, 16).Value = totalCutCost
    
    ThisWorkbook.Sheets("Sheet3").Cells(11, 3).Value = totalCutCost
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
