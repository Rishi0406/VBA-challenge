Attribute VB_Name = "Module3"

Sub alphabeticalC()
    
        ' define moderate result variables
    Dim ticker As String
    Dim year_open As Long
    Dim year_close As Long
    Dim yearly_range As Double
    Dim percent_change As Long
    Dim total_vol As Long
    Dim summary_table_row As Long
     
    ' define hard result variables
    Dim highest_TName As String
    Dim lowest_TName As String
    Dim greatvol_TName As String
    Dim highestvol_TName As String
    Dim highest_percent As Double
    Dim lowest_percent As Double
    Dim highest_vol As Long
    
    ' fixture of overflaw error
    On Error Resume Next
    
    ' print moderate result column names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Range"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' print hard result column & raw names
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increse"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    ' define their lowest value
    yearly_range = 0
    percent_change = 0
    total_vol = 0
    highest_percent = 0
    lowest_percent = 0
    highest_vol = 0
    summary_table_row = 2
    
    ' setting basic value before loop
    year_open = Cells(i, 3).Value
    
    ' loop to start
    For i = 2 To Rows.Count
        
    ' set condition to check and change ticker name and value and calculate yearly range and percent change
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ticker = Cells(i, 1).Value
            
            year_close = Cells(i, 6).Value
            yearly_range = year_close - year_open
            
            
            If year_open <> 0 Then
            
                percent_change = (yearly_range / year_open) * 100
            
            End If
            
    ' add values in columns
            total_vol = total_vol + Cells(i, 7).Value
            
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = yearly_range
            
    ' condition for conditional formatting
            If (yearly_range > 0) Then
            
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            
            ElseIf (yearly_range <= 0) Then
            
                Range("J" & summary_table_row).Interior.ColorIndex = 3
            
            End If
            
            Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")
            Range("L" & summary_table_row).Value = total_vol
    
    ' to continue adding by +1 value
            summary_table_row = summary_table_row + 1
            yearly_range = 0
            year_close = 0
            year_open = Cells(i + 1, 3).Value
    
    ' conditions to calculate hard result
            If (percent_change > highest_percent) Then
                
                highest_percent = percent_change
                highest_TName = ticker
            
            ElseIf (percent_change < lowest_percent) Then
                
                lowest_percent = percent_change
                lowest_TName = ticker
                
            End If
            
            If (total_vol > highest_vol) Then
                
                highest_vol = total_vol
                highestvol_TName = ticker
                
            End If
            
            percent_change = 0
            total_vol = 0
            
        Else
        
            total_vol = total_vol + Cells(i, 7).Value
            
        End If
    
    ' end loop
    Next i
            
    ' print hard result columns' value
    Range("Q2").Value = (CStr(highest_percent) & "%")
    Range("Q3").Value = (CStr(lowest_percent) & "%")
    Range("Q4").Value = highest_vol
    Range("P2").Value = highest_TName
    Range("P3").Value = lowest_TName
    Range("P4").Value = highestvol_TName
            
End Sub
