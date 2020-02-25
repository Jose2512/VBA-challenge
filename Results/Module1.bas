Attribute VB_Name = "Module1"
Sub stockdetails()
For Each ws In Worksheets
ws.Activate
' Adding headers for the new columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Volume"
' Declaring variables that need to be reset each i
Dim open_price As Double
Dim close_price As Double
Dim stock_volume As Double

'Counting last row
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
'Adding a counter to find the next row ONLY when the next ticker is changing
ins_ticker = 1
'First "for" for every row
For i = 2 To lastrow
    ' If previous ticker is different
    If Cells(i - 1, 1) <> Cells(i, 1) Then
        ' Reset variables
        stock_volume = 0
        open_price = 0
        close_price = 0
        'grab ticker name
        ticker = Cells(i, 1).Value
        'counter for displacing the ticker placement down
        ins_ticker = ins_ticker + 1
        'printing the ticker
        Cells(ins_ticker, 9) = ticker
        'storing open_price
        open_price = Cells(i, 3).Value
        'Storing stock_volumen
        stock_volume = Cells(i, 7)
    ' If previuos ticker is the same
    ElseIf Cells(i + 1, 1) = Cells(i, 1) Then
        'Add up the volume value
        stock_volume = stock_volume + Cells(i, 7)
    ' If next ticker is different
    ElseIf Cells(i + 1, 1) <> Cells(i, 1) Then
        'Store close_price
        close_price = Cells(i, 6).Value
        'Calculate diference
        Cells(ins_ticker, 10) = close_price - open_price
            'If to avoid division by "0"
            If open_price = 0 Then
            stock_variation = close_price
            Else
            stock_variation = (close_price - open_price) / open_price
            End If
        'Print the stock variation
        Cells(ins_ticker, 11) = stock_variation
        'Add the last volume value
        stock_volume = stock_volume + Cells(i, 7)
        'Print stock_volume and format it
        Cells(ins_ticker, 12) = stock_volume
        Cells(ins_ticker, 11).NumberFormat = "0.00%"
        'Formating for green and red
        If stock_variation < 0 Then
            Cells(ins_ticker, 10).Interior.ColorIndex = 3
        Else
            Cells(ins_ticker, 10).Interior.ColorIndex = 4
        End If
            
    End If
Next i

'Declaring variable for greatest increase, decrease and volume values and names
Dim G_Decrease As Double
G_Decrease = 0
Dim G_Increase As Double
G_Increase = 0
Dim G_Volumen As Double
G_Volumen = 0
Dim Name_G_Decrease As String
Dim Name_G_Increase As String
Dim Name_G_Volumen As String

'Adding headers
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest Percentage Increase"
Cells(3, 14).Value = "Greatest Percentage Decrease"
Cells(4, 14).Value = "Greatest Total Volume"


'counting last row in the new columns
lastrow2 = Cells(Rows.Count, "I").End(xlUp).Row

'Obtaining values and repleacing if necesary in the variables for each row
For j = 2 To lastrow2

    If Cells(j, 11) < G_Decrease Then
        G_Decrease = Cells(j, 11)
        Name_G_Decrease = Cells(j, 9)
    ElseIf Cells(j, 11) > G_Increase Then
        G_Increase = Cells(j, 11)
        Name_G_Increase = Cells(j, 9)
    End If
    If Cells(j, 12) > G_Volumen Then
        G_Volumen = Cells(j, 12)
        Name_G_Volumen = Cells(j, 9)
    End If
    
Next j

'Printing and formating the values
Cells(2, 15).Value = Name_G_Increase
Cells(2, 16).Value = G_Increase
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 15).Value = Name_G_Decrease
Cells(3, 16).Value = G_Decrease
Cells(3, 16).NumberFormat = "0.00%"
Cells(4, 15).Value = Name_G_Volumen
Cells(4, 16).Value = G_Volumen
Columns("A:Z").AutoFit
Next ws
End Sub

