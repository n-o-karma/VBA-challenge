VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarket()

'Loop through each worksheet
For Each ws In Worksheets

'Define a few variables for this section of code
Dim ticker As String
Dim change1 As Double
Dim change2 As Double
Dim vol As Double

vol = 0
tablerow = 2 'The table will begin with row 2 and iteratively crawl down
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Initialize the table contents
ws.Range("I" & 1).Value = "Ticker"
ws.Range("J" & 1).Value = "Yearly Change"
ws.Range("K" & 1).Value = "Percent Change"
ws.Range("L" & 1).Value = "Total Stock Volume"

'Loop through the ticker row
For i = 2 To lastrow
    
    'If there is a new ticker in the next row,
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        change2 = ws.Cells(i, 6).Value  'Update the closing value
        ticker = ws.Cells(i, 1).Value       'Update the ticker name
        vol = vol + ws.Cells(i, 7).Value    'Add to the total volume count
        ws.Range("I" & tablerow).Value = ticker 'Print the ticker found before the change
        ws.Range("J" & tablerow).Value = FormatNumber((change2 - change1), 2) 'Along with the yearly change, formatting the number to have 2 decimal places
        
        If (change2 - change1) < 0 Then
            ws.Range("J" & tablerow).Interior.ColorIndex = 3 'Formatting negative yearly change
        ElseIf (change2 - change1) > 0 Then
             ws.Range("J" & tablerow).Interior.ColorIndex = 4 'Formatting positive change
        Else
            ws.Range("J" & tablerow).Interior.ColorIndex = 15 'Formatting no change
        End If
        
        'Print the total stock volume;
        ws.Range("L" & tablerow).Value = vol
        'Calculate and print the percent change, formatting the number as a percent
        ws.Range("K" & tablerow).Value = Format((change2 - change1) / change1, "Percent")
    
        tablerow = tablerow + 1 'Move into the next table row,
        vol = 0                 'reset the stock volume
        change2 = 0             'along with the stored closing
        change1 = 0             'and opening values
    
    'If the value /above/ the cell in question is different, we can extract the next opening value
    ElseIf ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
    change1 = ws.Cells(i, 3).Value
    
    'Otherwise, add onto the total volume
    Else
    vol = vol + ws.Cells(i, 7).Value
    
    End If
    
Next i

'Define more variables for this section of code
Dim tempticker As String
Dim tempticker2 As String
Dim tempticker3 As String
Dim tempmax As Double
Dim tempmin As Double
Dim maxvol As Double

'Initialize the new table contents
ws.Range("P" & 1).Value = "Ticker"
ws.Range("Q" & 1).Value = "Value"
ws.Range("O" & 2).Value = "Greatest % Increase"
ws.Range("O" & 3).Value = "Greatest % Decrease"
ws.Range("O" & 4).Value = "Greatest Total Volume"

'Working with the percent change column now (column 11). Not stored as a variable this time since it doesnt change
lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Initialize the starting values
tempmax = ws.Cells(2, 11).Value 'Use the first value as the max
tempmin = tempmax               'and as the min, since none others have been considered
tempticker = ws.Cells(2, 9).Value 'The corresponding first ticker's name
tempticker2 = tempticker            'These will be the tickers for the second
tempticker3 = tempticker             'and third values in the new table, initialized to be the same
maxvol = ws.Cells(2, 12).Value  'Initialize the greatest volume for the corresponding first ticker

'Looping through the Percent Change column
For j = 2 To lastrow
    If tempmax < ws.Cells(j, 11).Value Then     'If we find a new greatest % increase,
        tempmax = ws.Cells(j, 11).Value         'Update the value
        tempticker = ws.Cells(j, 9).Value       'and the ticker name
    Else
        End If
        
    If tempmin > ws.Cells(j, 11).Value Then     'If we find a new greatest % decrease,
        tempmin = ws.Cells(j, 11).Value         'Update the value
        tempticker2 = ws.Cells(j, 9).Value      'and the ticker name
    Else
        End If
    If maxvol < ws.Cells(j, 12).Value Then      'If we find a new greatest total volume,
        maxvol = ws.Cells(j, 12).Value          'Update the value
        tempticker3 = ws.Cells(j, 9).Value      'and the ticker name
    Else
        End If
Next j

'Populating and formatting the new table
ws.Range("Q" & 2).Value = Format(tempmax, "Percent")
ws.Range("P" & 2).Value = tempticker
ws.Range("Q" & 3).Value = Format(tempmin, "Percent")
ws.Range("P" & 3).Value = tempticker2
ws.Range("Q" & 4).Value = Format(maxvol, "Scientific")
ws.Range("P" & 4).Value = tempticker3
ws.Range("I:Q").EntireColumn.AutoFit

'Move into the next worksheet

Next ws

End Sub
