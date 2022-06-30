Attribute VB_Name = "Module1"
Option Explicit

Sub stock()

' setup script for loop to run through each worksheet, a and b used for this.

    Dim a As Integer, b As Integer
    
    b = Application.Worksheets.Count
    
    ' For loop to run through all the worksheets (contains all code that updates worksheets)
    
    For a = 1 To b
        
        ' activate current worksheet
        
        Worksheets(a).Activate
        
        ' setup variable to determine last row of each spreadsheet
        
        Dim LR As Double
        
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' fill in cells/columns for organized data based on subsequent loops
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Total Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' dim variables for running through ticker (i) and entering new table (ticker/total change/percent change/total volume) (j)
        
        Dim i As Double, j As Double
        
        ' j will be the row of the new table that will be created by each unique ticker/stock
        i = 2
        j = 2
        
        ' include a ticker/value that will reset with every new ticker/stock encountered. start at day 0
        
        Dim day_of_year As Integer
        
        day_of_year = 0
                
           ' iterate over "A" column
           
        For i = 2 To LR
        
            ' if the current stock text matches the next cell. The goal is to figure out when the ticker column changes names indicating a new stock
            
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            
                 ' then nothing happens regarding calculating a new stock or subsequent calcs
                
                j = j
                
                'advance the day_of_year ticker to indicate we are still reviewing the same stock and have not reached the end of the year yet
                
                day_of_year = day_of_year + 1
                
            ' else the name of the ticker has changed indicating a new stock ticker and indicating time to calculate
            
            Else
                                                                          
                ' input ticker into ticker column by unique values
                
                Cells(j, 9).Value = Cells(i, 1).Value
                
                ' calculate total change from last day close to first day open
                
                Cells(j, 10).Value = Cells(i, 6).Value - Cells(i - day_of_year, 3).Value
                
                If Cells(j, 10) > 0 Then
       
                    Cells(j, 10).Interior.Color = RGB(0, 255, 0)
                    
                ElseIf Cells(j, 10) = 0 Then
                
                    Cells(j, 10).Interior.Color = RGB(255, 255, 255)
                
                Else

                Cells(j, 10).Interior.Color = RGB(255, 0, 0)

                End If
                                             
                ' calculate percent change over year
                
                Cells(j, 11).Value = ((Cells(i, 6).Value - Cells(i - day_of_year, 3).Value) / Cells(i - day_of_year, 3).Value)
                
                If Cells(j, 11) > 0 Then
       
                    Cells(j, 11).Interior.Color = RGB(0, 255, 0)
                    
                ElseIf Cells(j, 11) = 0 Then
                
                    Cells(j, 11).Interior.Color = RGB(255, 255, 255)
                
                Else

                    Cells(j, 11).Interior.Color = RGB(255, 0, 0)

                End If
                
                'format percent change to percentage style/format
                
                Cells(j, 11).NumberFormat = "0.00%"
                
                ' calculate total volume = sum of volume for every day market open
                
                Cells(j, 12).Value = Application.WorksheetFunction.Sum(Range("G" & i - day_of_year, "G" & i))
                
                ' j variable increases by 1 to put table into next ticker row
                
                j = j + 1
                
                 ' reset day of year for next ticker in list
                
                day_of_year = 0
            
            End If
        
        Next i
        
        ' ----------------------------------------------------
        
        'Bonus question
        
        Dim n As Integer
        ' set cells being calculated at zero as a baseline comparison for For loop and if statements below
        Range("Q2").Value = 0
        Range("Q3").Value = 0
        Range("Q4").Value = 0
        
        'loop for going through the new table created
        
        For n = 2 To 3001
        
            'Greatest percent increase if then statement
            
            If Cells(n, 11).Value > Cells(n + 1, 11).Value And Cells(n, 11) > Range("Q2").Value Then
            
            Range("Q2").Value = Cells(n, 11).Value
            
            Else
            
            Range("Q2").Value = Range("Q2").Value
                        
            End If
            
            'format percent change to percentage style/format
            Range("Q2").NumberFormat = "0.00%"
            
            'Greatest percent decrease if then statement
            
            If Cells(n, 11).Value < Cells(n + 1, 11).Value And Cells(n, 11) < Range("Q3").Value Then
            
            Range("Q3").Value = Cells(n, 11).Value
            
            End If
            
            'format percent change to percentage style/format
            Range("Q3").NumberFormat = "0.00%"
            
            'Greatest total volume if then statement
            
            If Cells(n, 12).Value > Cells(n + 1, 12).Value And Cells(n, 12) > Range("Q4").Value Then
            
            Range("Q4").Value = Cells(n, 12).Value
            
            Else
            
            Range("Q4").Value = Range("Q4").Value
            
            End If
            
            ' backtracking to find the ticker code associated with greatest percent increase
            
            If Range("Q2").Value = Cells(n, 11).Value Then
            
                Range("P2").Value = Cells(n, 9).Value
                
            End If
            
            ' backtracking to find the ticker code associated with greatest percent decrease
            
            If Range("Q3").Value = Cells(n, 11).Value Then
            
                Range("P3").Value = Cells(n, 9).Value
                
            End If
            
            ' backtracking to find the ticker code associated with greatest total volume
            
            If Range("Q4").Value = Cells(n, 12).Value Then
            
                Range("P4").Value = Cells(n, 9).Value
                
            End If
            
        Next n
        
        Columns("I:R").AutoFit
    Next a
    
            
    End Sub
    
