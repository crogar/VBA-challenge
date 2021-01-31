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
Set wb = Workbooks("Multiple_year_stock_data.xlsx") 'Setting this as our workbook

first_opening = True  'Indicates that this would be the first "open" row when detecting a new ticker'
counter = 2 'first Ticker will be printed in row 2'

End Sub
