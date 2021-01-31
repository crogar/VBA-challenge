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
  msgbox(ws.name)
next ws

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         Subroutines and Functions               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Function that finds and return the last non-blank cell in column A(1)
Private Function get_lr(ws_ As Workbook, SN_, col) As Long
'SN_ represents the sheet number and col represents the column to use as
'reference
Dim lRow As Long
    lRow = ws_.Cells(Rows.Count, col).End(xlUp).row
    get_lr = lRow
End Function