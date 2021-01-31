Sub main()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                Variable Declaration             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim wb As Workbook
Dim counter As Long      'Varible used to know last row where to prin next ticker
Dim previous_ctr As Double      'Varible used to count <ticker>s
Dim temp_open As Double  'Variable that temporarily will hold the Open price

Dim first_opening As Boolean 'Variable used to help us know when an open occurs in a specific company
Dim typt(3) As Double    ' typt = "Yearly Change, Percent Change, Total Stock volume"
' typt(0) -> Yearly Change, typt(1) -> Percent Change, typt(2) -> Total Stock volume
Set wb = Workbooks("alphabetical_testing.xlsx") 'Setting this as our workbook

first_opening = True  'Indicates that this would be the first "open" row when detecting a new ticker'
counter = 2 'first Ticker will be printed in row 2'

For Each ws In wb.Sheets  'i is the variable used to iterate through sheets
    ws.Cells.ClearFormats
    
    For row = 2 To get_lr(ws, 1)
         typt(2) = typt(2) + ws.Cells(row, 7).Value2
        'Start volumes up until we find a new Ticker!
        If first_opening = True Then
           temp_open = ws.Cells(row, 3).value
           first_opening = False
        End If
        If ws.Cells(row + 1, 1).value <> ws.Cells(row, 1).value Then
            ' Setting the <Ticker>
            ws.Cells(counter, 9).value = ws.Cells(row, 1).value
            ' Setting the Yearly change total
            ' WHERE ws.Cells(Row, 6).Value represents <close>
            ws.Cells(counter, 10).value = ws.Cells(row, 6).value - temp_open
            ' temporarily storing Yearly change
            typt(0) = ws.Cells(row, 6).value - temp_open
            'Styling Yearly change cell
            If typt(0) < 0 Then
                ws.Cells(counter, 10).Interior.Color = vbRed
            Else
                ws.Cells(counter, 10).Interior.Color = vbGreen
            End If
            ' Setting The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            If typt(0) = 0 Or temp_open = 0 Then
                ws.Cells(counter, 11).value = 0
                typt(1) = 0
            Else
                ws.Cells(counter, 11).value = typt(0) / temp_open
                typt(1) = typt(0) / temp_open
            End If
            ' Setting The Total Stock Volume
            ws.Cells(counter, 12).value = typt(2)
            
                        '*Re-asigning variables*
            typt(2) = 0        'total stock volume reset to 0
            counter = counter + 1   'adding 1, next company
            first_opening = Not first_opening   'this will help resetting our
        End If
        
    Next row
    counter = 2
    Call bonus_analysis(ws)
    Call styling(ws)
Next ws

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         Subroutines and Functions               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Function that finds and return the last non-blank cell in column A(1)
Private Function get_lr(ws_, col) As Long
'SN_ represents the sheet number and col represents the column to use as
'reference
Dim lRow As Long
    lRow = ws_.Cells(Rows.Count, col).End(xlUp).row
    get_lr = lRow
End Function

'Subroutine that styles the worksheets
Sub styling(ws_)
    Dim lr As Integer: lr = get_lr(ws_, 9)
    ws_.Cells(1, 9).value = "Ticker"
    ws_.Cells(1, 10).value = "Yearly Change"
    ws_.Cells(1, 11).value = "Percent Change"
    ws_.Cells(1, 12).value = "Total Stock volume"
    ws_.Range("K2:K" & lr).NumberFormat = "0.00%"
    ws_.Columns("O").AutoFit
End Sub

'*******************************************************************
'                         Bonus                                   *
'*****************************************************************
'Subroutine that calculates "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
Sub bonus_analysis(ws_)
'obtaining the last non-blank cell on J column
Dim lr As Integer: lr = get_lr(ws_, 9)

ws_.Range("P1").value = "Ticker"
ws_.Range("Q1").value = "Value"
ws_.Range("O2").value = "Greatest % increase"
ws_.Range("O3").value = "Greatest % decrease"
ws_.Range("O4").value = "Greatest total volume"

'Formatting Values in the following range to be shown as Percentage
ws_.Range("Q2:Q3").NumberFormat = "0.00%"
'obtaining Greatest % Increase
Set Rng = ws_.Range("K2:K" & lr)
ws_.Range("Q2").value = Application.WorksheetFunction.Max(Rng)
'Setting the corresponding value for Ticker
ws_.Range("P2").value = find_ticker(ws_, "K", ws_.Range("Q2").Value2)
'obtaining Greatest % decrease
Set Rng = ws_.Range("K2:K" & lr)
ws_.Range("Q3").value = Application.WorksheetFunction.Min(Rng)
'Setting the corresponding ticker for Greatest % decrease
ws_.Range("P3").value = find_ticker(ws_, "K", ws_.Range("Q3").Value2)
'obtaining Greatest total volume
Set Rng = ws_.Range("L2:L" & lr)
ws_.Range("Q4").value = Application.WorksheetFunction.Max(Rng)
'Setting the corresponding ticker for Greatest % decrease
ws_.Range("P4").value = find_ticker(ws_, "L", ws_.Range("Q4").Value2)

End Sub

'This Function is used to find a specific value in a given range
'and it returns the row where that value was found, otherwise will
'show "Ticker not found" in a msgbox
Private Function find_ticker(ws_, col As String, value As Variant) As String
Dim row As Integer: row = get_lr(ws_, 9)
Dim L_row As Long
Set rgFound = ws_.Range(col & "2:" & col & row).Find(value)

If Not rgFound Is Nothing Then
    L_row = rgFound.row
    find_ticker = ws_.Range("I" & L_row).Value2
Else
    MsgBox "Ticker not found"
End If

End Function

