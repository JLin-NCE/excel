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
    Dim currentCutEnd As Double
    Dim cutArea As Double
    Dim sectionArea As Double
    Dim cutCost As Double
    Dim cutType As String
    Dim feeCalculation As String
    
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
    
    ' Get input values from Sheet3
    streetName = CStr(wsSheet3.Cells(3, 3).Value) ' C3
    startLocation = CStr(wsSheet3.Cells(4, 3).Value) ' C4
    endLocation = CStr(wsSheet3.Cells(5, 3).Value) ' C5
    cutLength = CDbl(wsSheet3.Cells(6, 3).Value) ' C6
    cutWidth = CDbl(wsSheet3.Cells(7, 3).Value) ' C7
    distanceFromPrevSection = CDbl(wsSheet3.Cells(8, 3).Value) ' C8
    anticipatedCutYear = CInt(wsSheet3.Cells(9, 3).Value) ' C9
    inflationRate = CDbl(wsSheet3.Cells(10, 3).Value) ' C10
    
    ' Calculate total cut start point
    totalCutStart = distanceFromPrevSection
    
    ' Output headers in the output sheet, starting from column E
    wsSheet3Output.Cells(1, 5).Value = "Street Name"
    wsSheet3Output.Cells(1, 6).Value = "From"
    wsSheet3Output.Cells(1, 7).Value = "To"
    wsSheet3Output.Cells(1, 8).Value = "Length"
    wsSheet3Output.Cells(1, 9).Value = "Width"
    wsSheet3Output.Cells(1, 10).Value = "Area"
    wsSheet3Output.Cells(1, 11).Value = "PCI"
    wsSheet3Output.Cells(1, 12).Value = "Functional Class"
    wsSheet3Output.Cells(1, 13).Value = "Small Cut Fee"
    wsSheet3Output.Cells(1, 14).Value = "Large Cut Fee"
    wsSheet3Output.Cells(1, 15).Value = "Section Start"
    wsSheet3Output.Cells(1, 16).Value = "Section End"
    wsSheet3Output.Cells(1, 17).Value = "Cut Type"
    wsSheet3Output.Cells(1, 18).Value = "Cut Cost"
    wsSheet3Output.Cells(1, 19).Value = "Cut Area"
    wsSheet3Output.Cells(1, 20).Value = "Fee Calculation"
    
    ' Find the start row in the PCI Report sheet
    For startRow = 2 To wsPCI.Cells(Rows.Count, 3).End(xlUp).row
        If CStr(wsPCI.Cells(startRow, 3).Value) = streetName And CStr(wsPCI.Cells(startRow, 4).Value) = startLocation Then
            foundStart = True
            Exit For
        End If
    Next startRow
    
    If Not foundStart Then
        MsgBox "Beginning location not found for Street Name: " & streetName & ", From: " & startLocation, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Iterate to find the end location
    currentRow = startRow
    Do While currentRow <= wsPCI.Cells(Rows.Count, 3).End(xlUp).row
        If CStr(wsPCI.Cells(currentRow, 3).Value) = streetName And CStr(wsPCI.Cells(currentRow, 5).Value) = endLocation Then
            endRow = currentRow
            foundEnd = True
            Exit Do
        End If
        currentRow = currentRow + 1
    Loop
    
    If Not foundEnd Then
        MsgBox "Ending location not found for Street Name: " & streetName & ", To: " & endLocation & vbCrLf & _
               "Starting search from row: " & startRow, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Initialize remaining cut length
    remainingCutLength = cutLength
    currentCutStart = totalCutStart
    
    ' Extract and output data directly in the output sheet from column E onwards
    For row = startRow To endRow
        sectionLength = wsPCI.Cells(row, 10).Value
        sectionWidth = wsPCI.Cells(row, 11).Value
        sectionStart = currentCutStart
        sectionEnd = sectionStart + sectionLength
        
        ' Check if the section falls within the remaining cut length
        If sectionEnd > totalCutStart + cutLength Then
            sectionEnd = totalCutStart + cutLength
            sectionLength = sectionEnd - sectionStart
        End If
        
        ' Output section dimensions
        wsSheet3Output.Cells(outputRow, 5).Value = wsPCI.Cells(row, 3).Value ' Street Name
        wsSheet3Output.Cells(outputRow, 6).Value = wsPCI.Cells(row, 4).Value ' From
        wsSheet3Output.Cells(outputRow, 7).Value = wsPCI.Cells(row, 5).Value ' To
        wsSheet3Output.Cells(outputRow, 8).Value = sectionLength ' Length
        wsSheet3Output.Cells(outputRow, 9).Value = sectionWidth ' Width
        wsSheet3Output.Cells(outputRow, 10).Value = sectionLength * sectionWidth ' Area
        wsSheet3Output.Cells(outputRow, 11).Value = wsPCI.Cells(row, 14).Value ' PCI
        
        ' Calculate the unit cost based on the fee table
        functionalClass = wsPCI.Cells(row, 8).Value ' Use Rank column (H)
        pci = wsPCI.Cells(row, 14).Value
        sectionArea = sectionLength * sectionWidth
        cutArea = sectionLength * cutWidth
        
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
        If cutArea < 0.1 * sectionArea Then
            cutType = "Small Cut"
            cutCost = cutArea * smallCutFee
            feeCalculation = cutArea & " * " & smallCutFee
        Else
            cutType = "Large Cut"
            cutCost = cutArea * largeCutFee
            feeCalculation = cutArea & " * " & largeCutFee
        End If
        
        wsSheet3Output.Cells(outputRow, 12).Value = functionalClassFull
       
        wsSheet3Output.Cells(outputRow, 12).Value = functionalClassFull
        wsSheet3Output.Cells(outputRow, 13).Value = smallCutFee
        wsSheet3Output.Cells(outputRow, 14).Value = largeCutFee
        wsSheet3Output.Cells(outputRow, 15).Value = sectionStart
        wsSheet3Output.Cells(outputRow, 16).Value = sectionEnd
        wsSheet3Output.Cells(outputRow, 17).Value = cutType
        wsSheet3Output.Cells(outputRow, 18).Value = cutCost
        wsSheet3Output.Cells(outputRow, 19).Value = cutArea
        wsSheet3Output.Cells(outputRow, 20).Value = feeCalculation
        
        ' Update remaining cut length and current cut start
        remainingCutLength = remainingCutLength - sectionLength
        currentCutStart = sectionEnd
        
        ' If the remaining cut length is less than or equal to zero, exit the loop
        If remainingCutLength <= 0 Then
            Exit For
        End If
        
        outputRow = outputRow + 1
    Next row
    
    MsgBox "Data extraction complete from row " & startRow & " to row " & endRow, vbInformation, "Success"
End Sub




