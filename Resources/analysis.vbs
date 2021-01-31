sub main()
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

For each ws in wb.Sheets  'i is the variable used to iterate through sheets
	ws.Cells.ClearFormats
	
	For row = 2 To get_lr(ws,1)
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
			
			
			            '*Re-asigning variables*
            typt(2) = 0        'total stock volume reset to 0
            counter = counter + 1   'adding 1, next company
            first_opening = Not first_opening   'this will help resetting our
		endif
		
	next row
	counter = 2
next ws

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