Option Explicit

' Type to hold input data
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

' Main procedure
Public Sub GatherAssociatedRows()
    Dim wsSheet3 As Worksheet
    Dim wsSheet3Output As Worksheet
    Dim wsPCI As Worksheet
    Dim inputData As InputDataType
    Dim startRow As Long, endRow As Long
    
    ' Set worksheets
    Set wsSheet3 = ThisWorkbook.Sheets("Sheet3")
    Set wsPCI = ThisWorkbook.Sheets("Covina PCI Report")
    Set wsSheet3Output = CreateOrClearOutputSheet
    
    ' Get input data
    inputData = GetInputData(wsSheet3)
    
    ' Find start and end rows
    If Not FindStartAndEndRows(wsPCI, inputData, startRow, endRow) Then Exit Sub
    
    ' Process data and output results
    ProcessAndOutputData wsSheet3Output, wsPCI, inputData, startRow, endRow
    
    MsgBox "Data extraction complete from row " & startRow & " to row " & endRow, vbInformation, "Success"
End Sub

' Function to create or clear the output sheet
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
    
    ' Set headers
    With ws
        .Cells(1, 1).Value = "Street Name"
        .Cells(1, 2).Value = "From"
        .Cells(1, 3).Value = "To"
        .Cells(1, 4).Value = "Section Start"
        .Cells(1, 5).Value = "Section End"
        .Cells(1, 6).Value = "Length"
        .Cells(1, 7).Value = "Width"
        .Cells(1, 8).Value = "Area"
        .Cells(1, 9).Value = "PCI"
        .Cells(1, 10).Value = "Functional Class"
        .Cells(1, 11).Value = "Cut Type"
        .Cells(1, 12).Value = "Cut Area"
        .Cells(1, 13).Value = "Small Cut Fee"
        .Cells(1, 14).Value = "Large Cut Fee"
        .Cells(1, 15).Value = "Fee Calculation"
        .Cells(1, 16).Value = "Cut Cost"
    End With
    
    Set CreateOrClearOutputSheet = ws
End Function

' Function to get input data from Sheet3
Private Function GetInputData(ws As Worksheet) As InputDataType
    Dim data As InputDataType
    
    With ws
        data.streetName = CStr(.Cells(3, 3).Value)
        data.startLocation = CStr(.Cells(4, 3).Value)
        data.endLocation = CStr(.Cells(5, 3).Value)
        data.cutLength = Round(CDbl(.Cells(6, 3).Value), 2)
        data.cutWidth = Round(CDbl(.Cells(7, 3).Value), 2)
        data.distanceFromPrevSection = Round(CDbl(.Cells(8, 3).Value), 2)
        data.anticipatedCutYear = CInt(.Cells(9, 3).Value)
        data.inflationRate = Round(CDbl(.Cells(10, 3).Value), 2)
    End With
    
    GetInputData = data
End Function

' Function to find start and end rows
Private Function FindStartAndEndRows(ws As Worksheet, inputData As InputDataType, ByRef startRow As Long, ByRef endRow As Long) As Boolean
    Dim currentRow As Long, lastRow As Long
    Dim foundStart As Boolean, foundEnd As Boolean
    Dim disruptiveEndLocation As String
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row
    
    ' Find start row
    For startRow = 2 To lastRow
        If CStr(ws.Cells(startRow, 3).Value) = inputData.streetName And CStr(ws.Cells(startRow, 4).Value) = inputData.startLocation Then
            foundStart = True
            Exit For
        End If
    Next startRow
    
    If Not foundStart Then
        MsgBox "Beginning location not found for Street Name: " & inputData.streetName & ", From: " & inputData.startLocation, vbExclamation, "Error"
        FindStartAndEndRows = False
        Exit Function
    End If
    
    ' Find end row
    For currentRow = startRow To lastRow
        If CStr(ws.Cells(currentRow, 3).Value) = inputData.streetName Then
            If CStr(ws.Cells(currentRow, 5).Value) = "END" Then
                If Not foundEnd Then
                    disruptiveEndLocation = CStr(ws.Cells(currentRow, 4).Value)
                    MsgBox "Error: Reached end of street before finding specified end location." & vbNewLine & _
                           "Street Name: " & inputData.streetName & vbNewLine & _
                           "Specified End Location: " & inputData.endLocation & vbNewLine & _
                           "Disruptive End Location: " & disruptiveEndLocation & " (END)", vbExclamation, "Error"
                    FindStartAndEndRows = False
                    Exit Function
                End If
            ElseIf CStr(ws.Cells(currentRow, 5).Value) = inputData.endLocation Then
                endRow = currentRow
                foundEnd = True
                Exit For
            End If
        Else
            MsgBox "Error: End location not found before street name changed." & vbNewLine & _
                   "Street Name: " & inputData.streetName & vbNewLine & _
                   "Specified End Location: " & inputData.endLocation, vbExclamation, "Error"
            FindStartAndEndRows = False
            Exit Function
        End If
    Next currentRow
    
    If Not foundEnd Then
        MsgBox "Ending location not found for Street Name: " & inputData.streetName & ", To: " & inputData.endLocation, vbExclamation, "Error"
        FindStartAndEndRows = False
        Exit Function
    End If
    
    FindStartAndEndRows = True
End Function

' Procedure to process data and output results
Private Sub ProcessAndOutputData(wsOutput As Worksheet, wsPCI As Worksheet, inputData As InputDataType, startRow As Long, endRow As Long)
    Dim outputRow As Long, row As Long
    Dim remainingCutLength As Double, currentCutStart As Double, totalCutCost As Double
    Dim sectionLength As Double, sectionWidth As Double, sectionArea As Double
    Dim sectionStart As Double, sectionEnd As Double
    Dim functionalClass As String, functionalClassFull As String, pci As Double
    Dim smallCutFee As Double, largeCutFee As Double
    Dim cutArea As Double, cutCost As Double, cutType As String, feeCalculation As String
    Dim maxAllowedWidth As Double
    
    outputRow = 2
    remainingCutLength = inputData.cutLength
    currentCutStart = inputData.distanceFromPrevSection
    totalCutCost = 0
    maxAllowedWidth = 100 ' Set a reasonable maximum width (adjust as needed)
    
    ' Check for negative cut length
    If inputData.cutLength <= 0 Then
        MsgBox "Error: Cut length must be positive. Current value: " & inputData.cutLength & vbNewLine & _
               "Street Name: " & inputData.streetName & vbNewLine & _
               "Beginning Location: " & inputData.startLocation, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Check for negative cut width
    If inputData.cutWidth <= 0 Then
        MsgBox "Error: Cut width must be positive. Current value: " & inputData.cutWidth & vbNewLine & _
               "Street Name: " & inputData.streetName & vbNewLine & _
               "Beginning Location: " & inputData.startLocation, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Check if distance from previous cross section exceeds section length
    sectionLength = Round(wsPCI.Cells(startRow, 10).Value, 2)
    If inputData.distanceFromPrevSection > sectionLength Then
        MsgBox "Error: Distance from previous cross section (" & inputData.distanceFromPrevSection & ") " & _
               "exceeds the length of the beginning section (" & sectionLength & ")." & vbNewLine & _
               "Street Name: " & inputData.streetName & vbNewLine & _
               "Beginning Location: " & inputData.startLocation, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Check if cut width is unreasonably large
    If inputData.cutWidth > maxAllowedWidth Then
        MsgBox "Error: Cut width (" & inputData.cutWidth & ") " & _
               "exceeds the maximum allowed width (" & maxAllowedWidth & ")." & vbNewLine & _
               "Street Name: " & inputData.streetName & vbNewLine & _
               "Beginning Location: " & inputData.startLocation, vbExclamation, "Error"
        Exit Sub
    End If
    
    For row = startRow To endRow
        sectionLength = Round(wsPCI.Cells(row, 10).Value, 2)
        sectionWidth = Round(wsPCI.Cells(row, 11).Value, 2)
        sectionArea = Round(wsPCI.Cells(row, 12).Value, 2)
        
        ' Calculate section start and end
        If row = startRow Then
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
        
        ' Calculate fees and costs
        functionalClass = wsPCI.Cells(row, 8).Value
        pci = Round(wsPCI.Cells(row, 14).Value, 2)
        cutArea = Round(sectionLength * inputData.cutWidth, 2)
        
        DetermineFees functionalClass, pci, functionalClassFull, smallCutFee, largeCutFee
        DetermineCutTypeAndCost cutArea, sectionArea, smallCutFee, largeCutFee, cutType, cutCost, feeCalculation
        
        ' Output section information
        OutputSectionInfo wsOutput, outputRow, wsPCI, row, sectionStart, sectionEnd, sectionLength, sectionWidth, pci, _
                         functionalClassFull, cutType, cutArea, smallCutFee, largeCutFee, feeCalculation, cutCost
        
        ' Update variables
        remainingCutLength = Round(remainingCutLength - sectionLength, 2)
        currentCutStart = Round(sectionEnd, 2)
        totalCutCost = Round(totalCutCost + cutCost, 2)
        
        If remainingCutLength <= 0 Then Exit For
        
        outputRow = outputRow + 1
    Next row
    
    ' Output total cut cost
    outputRow = outputRow + 1
    wsOutput.Cells(outputRow, 1).Value = "Total Cut Cost"
    wsOutput.Cells(outputRow, 16).Value = totalCutCost
    
    ' Output total cut cost to Sheet3
    ThisWorkbook.Sheets("Sheet3").Cells(11, 3).Value = totalCutCost
End Sub


' Procedure to determine fees based on functional class and PCI
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

' Procedure to determine cut type and cost
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

' Procedure to output section information
Private Sub OutputSectionInfo(ws As Worksheet, row As Long, wsPCI As Worksheet, pciRow As Long, sectionStart As Double, sectionEnd As Double, _
                              sectionLength As Double, sectionWidth As Double, pci As Double, functionalClassFull As String, cutType As String, _
                              cutArea As Double, smallCutFee As Double, largeCutFee As Double, feeCalculation As String, cutCost As Double)
    With ws
        .Cells(row, 1).Value = wsPCI.Cells(pciRow, 3).Value ' Street Name
        .Cells(row, 2).Value = wsPCI.Cells(pciRow, 4).Value ' From
        .Cells(row, 3).Value = wsPCI.Cells(pciRow, 5).Value ' To
        .Cells(row, 4).Value = Round(sectionStart, 2) ' Section Start
        .Cells(row, 5).Value = Round(sectionEnd, 2) ' Section End
        .Cells(row, 6).Value = Round(sectionLength, 2) ' Length
        .Cells(row, 7).Value = Round(sectionWidth, 2) ' Width
        .Cells(row, 8).Value = Round(sectionLength * sectionWidth, 2) ' Area
        .Cells(row, 9).Value = Round(pci, 2) ' PCI
        .Cells(row, 10).Value = functionalClassFull ' Functional Class
        .Cells(row, 11).Value = cutType ' Cut Type
        .Cells(row, 12).Value = Round(cutArea, 2) ' Cut Area
        .Cells(row, 13).Value = Round(smallCutFee, 2) ' Small Cut Fee
        .Cells(row, 14).Value = Round(largeCutFee, 2) ' Large Cut Fee
        .Cells(row, 15).Value = feeCalculation ' Fee Calculation
        .Cells(row, 16).Value = Round(cutCost, 2) ' Cut Cost
    End With
End Sub

