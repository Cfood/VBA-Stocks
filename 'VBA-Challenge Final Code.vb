'VBA-Challenge

Sub Stonks()

Dim wsCounter As Integer

Dim goToOne As Integer
'
' go to first ws
goToOne = (ActiveSheet.Index) + 1
Worksheets(goToOne - ActiveSheet.Index).Select


Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets

wsCounter = wsCounter + 1
'determine range of i
    Dim row_num As Long
       'Range("A2").Select
        'Range(Selection, Selection.End(xlDown)).Select
        'row_num = Selection.Rows.count + 1
        
        row_num = Cells(Rows.Count, 1).End(xlUp).Row
        
    
'Dim variables and extract ticker
 
        Dim ticker As String
        Dim yearly_volume As Double
        Dim row_counter As Double
        Dim open_price As Double
        Dim closing_price As Double
        Dim price_change As Double
        Dim percent_change As Double
        
         
         'Label new Columns
            Cells(1, 12).Value = "Tot Vol"
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yrly Change"
            Cells(1, 11).Value = "% Change"
            Cells(2, 14).Value = "Greatest Annual % Increase"
            Cells(3, 14).Value = "Greatest Annual % Decrease"
            Cells(4, 14).Value = "Greatest Annual Total Volume"
            Cells(1, 15).Value = "Ticker"
            Cells(1, 16).Value = "Value"
            Columns("N:N").EntireColumn.AutoFit
            
            
            Dim counter As Long
            counter = 0
            row_counter = 2
            
            
    'Loop thru stocks and determine initial values
   
    
    For i = 2 To row_num
            ticker = Cells(i, 1).Value
            
            
     
     'Aggregate yrly Volume per ticker
            If ticker = Cells(i + 1, 1).Value Then
                   yearly_volume = Cells(i, 7).Value + yearly_volume
            'ticker ends, calculate volume, price change, percent change
            ElseIf ticker <> Cells(i + 1, 1).Value Then
                    yearly_volume = Cells(i, 7).Value + yearly_volume
                    counter = i + 2 + counter
                    'Print values to new columns
                    Cells(row_counter, 12).Value = yearly_volume
                    Cells(row_counter, 9).Value = ticker
                    
                    Dim x As Double
                    x = counter - i
                    open_price = Cells(x, 3).Value
                    closing_price = Cells(i, 6).Value
                    price_change = closing_price - open_price
                    
                    'Evade "DIV/0" error
                    If open_price <> 0 Then
                    percent_change = (closing_price - open_price) / open_price
                    Else: percent_change = 0
                    End If
                    
            
            'Print price and percent change to cells
                    Cells(row_counter, 11).Value = Format(percent_change, "Percent")
                    Cells(row_counter, 10).Value = price_change
                    
                        
                    'Color Yrly change based on value
                        If Cells(row_counter, 10).Value < 0 Then
                            Cells(row_counter, 10).Interior.ColorIndex = 3
                        ElseIf Cells(row_counter, 10).Value > 0 Then
                            Cells(row_counter, 10).Interior.ColorIndex = 4
                            
                    
                        End If
                     
                    
                    'reset values for the next ticker
                    yearly_volume = 0
                    row_counter = row_counter + 1
                    counter = i - 1
                    
            End If
            Next i

    'Most extreme stock performances
    
    Dim increase As Double
    Dim increaseTicker As String
    Dim decrease As Double
    Dim decreaseTicker As String
    Dim greatestVolume As Double
    Dim volumeTicker As String
    
    For i = 2 To (row_counter - 1)
        If Cells(i, 11).Value < decrease Then
        decrease = Cells(i, 11).Value
        decreaseTicker = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value > increase Then
        increase = Cells(i, 11).Value
        increaseTicker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > greatestVolume Then
        greatestVolume = Cells(i, 12).Value
        volumeTicker = Cells(i, 9).Value
        End If
        
    Range("p3").Value = Format(decrease, "percent")
    Range("o3").Value = decreaseTicker
    Range("p2").Value = Format(increase, "percent")
    Range("o2").Value = increaseTicker
    Range("p4").Value = greatestVolume
    Range("o4").Value = volumeTicker
    
   

    Next i
    
    
    decrease = 0
    increase = 0
    greatestVolume = 0
    volumeTicker = " "
    



    If wsCounter <> ActiveWorkbook.Sheets.Count Then
        Worksheets(ActiveSheet.Index + 1).Select

    Else

        Worksheets((ActiveSheet.Index + 1) - wsCounter).Select
    End If
    
    

Next ws
        

End Sub





