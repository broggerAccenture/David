Sub GetErrors_Click()
    
   sheetName = "Errors"

    'Reset errors
    ThisWorkbook.Sheets(sheetName).Range("A1:C1000").Value = ""
    
    'Header
    ThisWorkbook.Sheets(sheetName).Range("A1").Value = "Form ID"
    ThisWorkbook.Sheets(sheetName).Range("B1").Value = "Testcase"
    ThisWorkbook.Sheets(sheetName).Range("C1").Value = "testSubject"
    
    'Set range of results
    Set Rng = ThisWorkbook.Sheets("FF - Test Design").Range("K1:K2000")
    i = 2
    
    For Each Row In Rng.Rows
        For Each cell In Row.Cells
            If cell.Value = "False" Then
            
                'Log testcases that failed
                form = ThisWorkbook.Sheets("FF - Test Design").Cells(Row.Row, 1).Text
                tcid = ThisWorkbook.Sheets("FF - Test Design").Cells(Row.Row, 3).Text
                Subject = ThisWorkbook.Sheets("FF - Test Design").Cells(Row.Row, 6).Text
                cell.Interior.Color = RGB(252, 228, 214) 'Color testcase result red

                ThisWorkbook.Sheets(sheetName).Cells(i, 1).Value = form
                ThisWorkbook.Sheets(sheetName).Cells(i, 2).Value = tcid
                ThisWorkbook.Sheets(sheetName).Cells(i, 3).Value = Subject
                
                i = i + 1
            ElseIf cell.Value <> "False" And ThisWorkbook.Sheets("FF - Test Design").Cells(Row.Row, 3).Value <> "" And IsNumeric(ThisWorkbook.Sheets("FF - Test Design").Cells(Row.Row, 3).Value) Then
                cell.Interior.Color = RGB(226, 239, 218)
            End If
        Next cell
    Next Row

End Sub