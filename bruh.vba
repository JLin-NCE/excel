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
