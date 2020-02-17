Sub StockMarket()

'Create Summary Table Headers

 Range("I1").Value = "Ticker Symbol"
 Range("J1").Value = "Yearly Change"
 Range("K1").Value = "Percent Change"
 Range("L1").Value = "Total Stock"
 
     'Set Headers for Second Table
    
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
    'Set values in Columns to autofit
    
       Columns("A:Q").AutoFit

'Set an initial variable for holding Ticker Symbol

    Dim Ticker_Symbol As String

' Set initial variable for holding Ticker Stock Volume

    Dim Stock_Volume As Double
    Stock_Volume = 0

' Set initial variable for holding StartRow

    Dim StartRow As Double
    StartRow = 0
    
' Set inital variable for holding Closing Value

    Dim Close_Value As Double
    Close_Value = 0
    
' Set initial Variable for holding Opening Value

    Dim Open_Value As Double
    Open_Value = 0
    
' Set initial Variable for holding Yearly Change

    Dim Yearly_Change As Double
    Yearly_Change = 0

' Set inital Variable for Percentage Change

    Dim Percent_Change As Double
    Percent_Change = 0

' ---------------------------------------------------------

'Keep track of the location for each Ticker in the Summary Table

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
        
' -----------------------------------------------------------

' Set variable for Loop

    Dim i As Long

' Set Last Row Count

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all Tickers

    For i = 2 To lastrow

' Check if we are still in the same Ticker

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
' ------------------------------------------------------------
    
    'Set Ticker Symbol
    
      Ticker_Symbol = Cells(i, 1).Value
    
    'Set Stock Volume
    
      Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
    'Print the Ticker to the Summary Table
    
      Range("I" & Summary_Table_Row).Value = Ticker_Symbol
    
    'Print the Stock Volume to the Summary Table
    
      Range("L" & Summary_Table_Row).Value = Stock_Volume
    
    
' -------------------------------------------------

        'Set year end Closing Value
        
          Close_Value = Cells(i, 6).Value
          
        'Set year Opening Value
        
          Open_Value = Cells(i - StartRow, 3).Value
        
        'Set Yearly Change Value
        
          Yearly_Change = Close_Value - Open_Value
        
        'Print the Yearly Change to the Summary Table
          
          Range("J" & Summary_Table_Row) = Yearly_Change

          
          
'-----------------------------------------------

' Set Percent Change and account for Values of "0"
    
        If Open_Value > 0 Then
    
            Percent_Change = Yearly_Change / Open_Value
    
                Else
    
            Open_Value = 0
    
         End If
    
    'Print the Percent Change to the Summary Table
          
    Range("K" & Summary_Table_Row) = Percent_Change
    
    ' Set Percent Change to Percent style
          
    Range("K" & Summary_Table_Row) = Format(Percent_Change, "00.00%")
          

' ---------------------------------------------------

' Set Color Conditional Formatting
    
    If Yearly_Change > 0 Then
    
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
    ElseIf Yearly_Change < 0 Then
    
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
    End If


' --------------------------------------------------------------
    
    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset Stock Total
    Stock_Volume = 0
    
    'Reset Start Row
    StartRow = 0

' ------------------------------------------------------------
    
' If the cell immediately following a row is the same Ticker..

Else

    'Add to the Stock Volume
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
    'Add to the StartRow Counter
    StartRow = StartRow + 1
        
    End If


  Next i


End Sub