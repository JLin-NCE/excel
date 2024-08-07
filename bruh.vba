Sub GatherAssociatedRows()
    Dim wsSheet3 As Worksheet
    Dim wsSheet3Output As Worksheet
    Dim wsPCI As Worksheet
    Dim streetName As String
    Dim startLocation As String
    Dim endLocation As String
    Dim cutLength As Double
    Dim cutWidth As Double
    Dim distanceFromPrevSection As Double
    Dim anticipatedCutYear As Integer
    Dim inflationRate As Double
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim foundStart As Boolean
    Dim foundEnd As Boolean
    Dim outputRow As Long
    Dim sectionLength As Double
    Dim sectionWidth As Double
    Dim remainingCutLength As Double
    Dim totalCutStart As Double
    Dim functionalClass As String
    Dim functionalClassFull As String
    Dim pci As Double
    Dim smallCutFee As Double
    Dim largeCutFee As Double
    Dim sectionStart As Double
    Dim sectionEnd As Double
    Dim currentCutStart As Double
    Dim cutArea As Double
    Dim sectionArea As Double
    Dim cutCost As Double
    Dim cutType As String
    Dim feeCalculation As String
    Dim totalCutCost As Double
    Dim isEndSegment As Boolean
    
    ' Set worksheets
    Set wsSheet3 = ThisWorkbook.Sheets("Sheet3")
    Set wsPCI = ThisWorkbook.Sheets("Covina PCI Report")
    
    ' Create or set the output worksheet
    On Error Resume Next
    Set wsSheet3Output = ThisWorkbook.Sheets("Sheet3 Output")
    If wsSheet3Output Is Nothing Then
        Set wsSheet3Output = ThisWorkbook.Sheets.Add(After:=wsSheet3)
        wsSheet3Output.Name = "Sheet3 Output"
    End If
    On Error GoTo 0
    
    ' Clear any existing content in the output sheet
    wsSheet3Output.Cells.Clear
    
    ' Initialize variables
    foundStart = False
    foundEnd = False
    outputRow = 2 ' Starting output row for results
    totalCutCost = 0 ' Initialize total cut cost
    
    ' Get input values from Sheet3
    streetName = CStr(wsSheet3.Cells(3, 3).Value) ' C3
    startLocation = CStr(wsSheet3.Cells(4, 3).Value) ' C4
    endLocation = CStr(wsSheet3.Cells(5, 3).Value) ' C5
    cutLength = Round(CDbl(wsSheet3.Cells(6, 3).Value), 2) ' C6
    cutWidth = Round(CDbl(wsSheet3.Cells(7, 3).Value), 2) ' C7
    distanceFromPrevSection = Round(CDbl(wsSheet3.Cells(8, 3).Value), 2) ' C8
    anticipatedCutYear = CInt(wsSheet3.Cells(9, 3).Value) ' C9
    inflationRate = Round(CDbl(wsSheet3.Cells(10, 3).Value), 2) ' C10
    
    ' Calculate total cut start point
    totalCutStart = Round(distanceFromPrevSection, 2)
    
    ' Output headers in the output sheet, starting from column A
    wsSheet3Output.Cells(1, 1).Value = "Street Name"
    wsSheet3Output.Cells(1, 2).Value = "From"
    wsSheet3Output.Cells(1, 3).Value = "To"
    wsSheet3Output.Cells(1, 4).Value = "Section Start"
    wsSheet3Output.Cells(1, 5).Value = "Section End"
    wsSheet3Output.Cells(1, 6).Value = "Length"
    wsSheet3Output.Cells(1, 7).Value = "Width"
    wsSheet3Output.Cells(1, 8).Value = "Area"
    wsSheet3Output.Cells(1, 9).Value = "PCI"
    wsSheet3Output.Cells(1, 10).Value = "Functional Class"
    wsSheet3Output.Cells(1, 11).Value = "Cut Type"
    wsSheet3Output.Cells(1, 12).Value = "Cut Area"
    wsSheet3Output.Cells(1, 13).Value = "Small Cut Fee"
    wsSheet3Output.Cells(1, 14).Value = "Large Cut Fee"
    wsSheet3Output.Cells(1, 15).Value = "Fee Calculation"
    wsSheet3Output.Cells(1, 16).Value = "Cut Cost"
    
    ' Find the start row in the PCI Report sheet
    For startRow = 2 To wsPCI.Cells(Rows.Count, 3).End(xlUp).Row
        If CStr(wsPCI.Cells(startRow, 3).Value) = streetName And CStr(wsPCI.Cells(startRow, 4).Value) = startLocation Then
            foundStart = True
            Exit For
        End If
    Next startRow
    
    If Not foundStart Then
        MsgBox "Beginning location not found for Street Name: " & streetName & ", From: " & startLocation, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Check if end location is "END"
    isEndSegment = (endLocation = "END")
    
    If isEndSegment Then
        endRow = startRow
        sectionLength = Round(wsPCI.Cells(startRow, 10).Value, 2) ' Length column
        If cutLength > sectionLength Then
            MsgBox "Cut length exceeds segment length for END segment. Adjusting cut length.", vbExclamation, "Warning"
            cutLength = sectionLength
        End If
        foundEnd = True
    Else
        ' Iterate to find the end location
        currentRow = startRow
        Do While currentRow <= wsPCI.Cells(Rows.Count, 3).End(xlUp).Row
            If CStr(wsPCI.Cells(currentRow, 3).Value) = streetName And CStr(wsPCI.Cells(currentRow, 5).Value) = endLocation Then
                endRow = currentRow
                foundEnd = True
                Exit Do
            End If
            currentRow = currentRow + 1
        Loop
    End If
    
    If Not foundEnd Then
        MsgBox "Ending location not found for Street Name: " & streetName & ", To: " & endLocation & vbCrLf & _
               "Starting search from row: " & startRow, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Initialize remaining cut length
    remainingCutLength = Round(cutLength, 2)
    currentCutStart = Round(totalCutStart, 2)
    
    ' Extract and output data directly in the output sheet from column A onwards
    For Row = startRow To endRow
        sectionLength = Round(wsPCI.Cells(Row, 10).Value, 2)
        sectionWidth = Round(wsPCI.Cells(Row, 11).Value, 2)
        sectionArea = Round(wsPCI.Cells(Row, 12).Value, 2)
        
        ' Calculate the length within the section
        If Row = startRow Then
            ' For the first section, calculate based on distance from start
            sectionStart = Round(distanceFromPrevSection, 2)
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
        
        ' Calculate the unit cost based on the fee table
        functionalClass = wsPCI.Cells(Row, 8).Value ' Use Rank column (H)
        pci = Round(wsPCI.Cells(Row, 14).Value, 2)
        cutArea = Round(sectionLength * cutWidth, 2)
        
        ' Determine fees based on functional class and PCI
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
        
        ' Determine cut type and cost
        If cutArea < Round(0.1 * sectionArea, 2) Then
            cutType = "Small Cut"
            cutCost = Round(cutArea * smallCutFee, 2)
            feeCalculation = cutArea & " * " & smallCutFee
        Else
            cutType = "Large Cut"
            cutCost = Round(cutArea * largeCutFee, 2)
            feeCalculation = cutArea & " * " & largeCutFee
        End If
        
        ' Output section information
        wsSheet3Output.Cells(outputRow, 1).Value = wsPCI.Cells(Row, 3).Value ' Street Name
        wsSheet3Output.Cells(outputRow, 2).Value = wsPCI.Cells(Row, 4).Value ' From
        wsSheet3Output.Cells(outputRow, 3).Value = wsPCI.Cells(Row, 5).Value ' To
        wsSheet3Output.Cells(outputRow, 4).Value = Round(sectionStart, 2) ' Section Start
        wsSheet3Output.Cells(outputRow, 5).Value = Round(sectionEnd, 2) ' Section End
        wsSheet3Output.Cells(outputRow, 6).Value = Round(sectionLength, 2) ' Length
        wsSheet3Output.Cells(outputRow, 7).Value = Round(sectionWidth, 2) ' Width
        wsSheet3Output.Cells(outputRow, 8).Value = Round(sectionLength * sectionWidth, 2) ' Area
        wsSheet3Output.Cells(outputRow, 9).Value = Round(pci, 2) ' PCI
        wsSheet3Output.Cells(outputRow, 10).Value = functionalClassFull ' Functional Class
        wsSheet3Output.Cells(outputRow, 11).Value = cutType ' Cut Type
        wsSheet3Output.Cells(outputRow, 12).Value = Round(cutArea, 2) ' Cut Area
        wsSheet3Output.Cells(outputRow, 13).Value = Round(smallCutFee, 2) ' Small Cut Fee
        wsSheet3Output.Cells(outputRow, 14).Value = Round(largeCutFee, 2) ' Large Cut Fee
        wsSheet3Output.Cells(outputRow, 15).Value = feeCalculation ' Fee Calculation
        wsSheet3Output.Cells(outputRow, 16).Value = Round(cutCost, 2) ' Cut Cost
        
        ' Update remaining cut length and current cut start
        remainingCutLength = Round(remainingCutLength - sectionLength, 2)
        currentCutStart = Round(sectionEnd, 2)
        
        ' Add to total cut cost
        totalCutCost = Round(totalCutCost + cutCost, 2)
        
        ' If the remaining cut length is less than or equal to zero, exit the loop
        If remainingCutLength <= 0 Then
            Exit For
        End If
        
        ' If it's an end segment, we're done after one iteration
        If isEndSegment Then
            Exit For
        End If
        
        outputRow = outputRow + 1
    Next Row
    
    ' Output total cut cost in a separate row
    outputRow = outputRow + 1
    wsSheet3Output.Cells(outputRow, 1).Value = "Total Cut Cost"
    wsSheet3Output.Cells(outputRow, 16).Value = totalCutCost
    
    ' Output total cut cost to cell C11 of Sheet3
    wsSheet3.Cells(11, 3).Value = totalCutCost
    
    MsgBox "Data extraction complete from row " & startRow & " to row " & endRow, vbInformation, "Success"
End Sub

