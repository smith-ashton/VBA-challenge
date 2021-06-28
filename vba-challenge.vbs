Attribute VB_Name = "Module1"
Sub wall_street()
    'INITIALIZE VARIABLES
    
    'initalize ticker and set it as the first name in the column for comparison sake
    Dim ticker As String
        ticker = Cells(2, 1).Value
    'initalize volume variable and set it to zero
    Dim volume As Double
       volume = 0
    'initalize open (first) and closing (last) variables to keep up with the opening and closing price of each ticker
    Dim first As Double
        first = Cells(2, 3)
    Dim last As Double
    'initialize change and percent variables to calculate later
    Dim change As Double
    Dim percent As Double
    'create lastrow variable
    Dim lastrow As Long
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'create summary chart
    Dim Summary_Row As Long
     Summary_Row = 1
     
     Cells(1, 9).Value = "Ticker"
     Cells(1, 10).Value = "Yearly Change"
     Cells(1, 11).Value = "Percent Change"
     Cells(1, 12).Value = "Total Volume"
     
    'FOR LOOP - Loop through data
    For i = 2 To lastrow
        'if same, add volume to total volume and keep ticker name the same
        If Cells(i, 1) = ticker Then
            volume = volume + Cells(i, 7).Value
        'if not the same...
        Else
            'set the previous closing value as the last value for that ticker
            last = Cells(i - 1, 6).Value
            'then calculate the yearly change and percent change
            change = last - first
            
            If first <> 0 Then
            percent = change / first
            Else
            percent = 0
            End If
            
            'go to next row of summary table
            Summary_Row = Summary_Row + 1
            'list the ticker, total volume, yearly change, and percent change to summary table
            Cells(Summary_Row, 9).Value = ticker
            Cells(Summary_Row, 10).Value = change
            Cells(Summary_Row, 11).Value = percent
            Cells(Summary_Row, 12).Value = volume
            'and go to the next line for the next ticker
            
            'reset the variables for the next ticker
            ticker = Cells(i, 1).Value
            volume = Cells(i, 7).Value
            first = Cells(i, 3).Value
            
        End If
    Next i

    'FOR LOOP - conditional coloring
    For i = 2 To Summary_Row
        'if negative (less than 0), color red
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        'if positive (greater than or equal to 0), color green
        Else
            Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
    
    '**BONUS CHALLENGE**
    'initialize variables and set them all to the first row in summary table
    Dim greatest_inc As Double
        greatest_inc = Cells(2, 11).Value
    Dim greatest_dec As Double
        greatest_dec = Cells(2, 11).Value
    Dim greatest_tot As Double
        greatest_tot = Cells(2, 12).Value
    Dim bonus_tic_inc As String
    Dim bonus_tic_tot As String
    
    'FOR LOOP - search summary table for greatest increase and decrease
    For i = 3 To Summary_Row
        'if the value it's on is greater than the current greatest_inc, it becomes greatest_inc
        'and the ticker that is on that row becomes the new bonus_tic_inc
        If Cells(i, 11).Value > greatest_inc Then
            greatest_inc = Cells(i, 11).Value
            bonus_tic_inc = Cells(i, 9).Value
        End If
        'if the value it's on is less than the current greatest_dec, it becomes the new greatest_dec
        'and the ticker that is on the row becomes the new bonus_tic_dec
        If Cells(i, 11).Value < greatest_dec Then
            greatest_dec = Cells(i, 11).Value
            bonus_tic_dec = Cells(i, 9).Value
        End If
    Next i
    'FOR LOOP - search summary table for greatest total
       For i = 3 To Summary_Row
        'if the value it's on is less than the current greatest_tot, it becomes new greatest_tot
        If Cells(i, 12).Value > greatest_tot Then
            greatest_tot = Cells(i, 12).Value
            bonus_tic_tot = Cells(i, 9).Value
        End If
    Next i
    
    'CREATE TABLE with final values
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 16).Value = bonus_tic_inc
    Cells(2, 17).Value = greatest_inc
    Cells(3, 16).Value = bonus_tic_dec
    Cells(3, 17).Value = greatest_dec
    Cells(4, 16).Value = bonus_tic_tot
    Cells(4, 17).Value = greatest_tot
    
    'format greatest % increase and decrease to be percentage
    Range("Q2:Q3").NumberFormat = "00.00%"
   'END BONUS
   
   'format percent change column to be percentage
   Range("K2:K" & Summary_Row).NumberFormat = "00.00%"

End Sub

