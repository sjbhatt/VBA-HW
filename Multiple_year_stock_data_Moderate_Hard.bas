Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis_Pt2()

    For Each file In Worksheets
    
        'Set variables for holding the Ticker Symbol Names
        Dim Stock_Name As String
        Dim Stock_Greatest_Percent_Increase As String
        Dim Stock_Greatest_Percent_Decrease As String
        Dim Stock_Greatest_Total_Volume As String
    
        ' Set an initial variable for holding the total per Ticker Symbol
        Dim Stock_Volume_Total As Double
        Stock_Volume_Total = 0

        ' Keep track of the location for each ticker name in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Keep track of range counter for the same ticker name data
        Dim counter As Integer
        counter = 0

    
        ' Add the Column Headers for Summary Table - 1
        file.Cells(1, 9).Value = "Ticker"
        file.Cells(1, 10).Value = "Yearly Change"
        file.Cells(1, 11).Value = "Percent Change"
        file.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Add the Headers for Summary Table - 2
        file.Cells(2, 14).Value = "Greatest % increase"
        file.Cells(3, 14).Value = "Greatest % Decrease"
        file.Cells(4, 14).Value = "Greatest total volume"
        file.Cells(1, 15).Value = "Ticker"
        file.Cells(1, 16).Value = "Value"

        ' Autofit to display data
        file.Columns("I:P").AutoFit
        
        Dim LastRow As Long
                
        ' Determine the Last Row
        LastRow = file.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all Ticker Symbols in the list
        For i = 2 To LastRow

            ' Check if we are still within the same Ticker Symbol Name, if it is not...
            If file.Cells(i + 1, 1).Value <> file.Cells(i, 1).Value Then

                ' Set the Ticker Symbol Name
                Stock_Name = file.Cells(i, 1).Value

                ' Add to the Stock_Volume_Total
                Stock_Volume_Total = Stock_Volume_Total + file.Cells(i, 7).Value

                ' Print the Ticker Symbol Name in the Summary Table
                file.Range("I" & Summary_Table_Row).Value = Stock_Name

                ' Print the Ticker Volume Total to the Summary Table
                file.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
                    
                ' Get the Stock Open and Close values
                Ticker_Close_Value = file.Cells(i, 6).Value
                Ticker_Open_Value = file.Cells(i - counter, 3).Value
                    
                ' Calculate the Yearly Change
                Change = Ticker_Close_Value - Ticker_Open_Value
                   
                ' For Divide by 0 errors
                If Ticker_Open_Value = 0 Then
                    Percent = 0
                Else
                    Percent = (Change / Ticker_Open_Value)
                End If
                    
                ' Populate the Summary Table
                file.Range("J" & Summary_Table_Row).Value = Change
                
                ' Formatting the Values in Summary Table
                If Change >= 0 Then
                    file.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    file.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                    
                file.Range("K" & Summary_Table_Row).Value = FormatPercent(Percent, 2)
                                        

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the Stock_Volume_Total
                Stock_Volume_Total = 0
                    
                ' Reset the Counter
                counter = 0

                ' If the cell immediately following a row is the same ticker symbol...
            Else

                ' Add to the Stock_Volume_Total
                Stock_Volume_Total = Stock_Volume_Total + file.Cells(i, 7).Value
                counter = counter + 1

            End If

        Next i
        
        ' Set variables for holding the values
        Dim Greatest_Percent_Increase As Double
        Dim Greatest_Percent_Decrease As Double
        Dim Greatest_Total_Volume As Double
       
        ' Initialize the Variables
        Greatest_Percent_Increase = file.Range("K2").Value
        Greatest_Percent_Decrease = file.Range("K2").Value
        Greatest_Total_Volume = file.Range("L2").Value
        
        ' Loop through the Summary Table 1 to find the max and min values
        For i = 2 To Summary_Table_Row
            
            If (Greatest_Percent_Increase < file.Range("K" & i).Value) Then
                Greatest_Percent_Increase = file.Range("K" & i).Value
                Stock_Greatest_Percent_Increase = file.Range("I" & i).Value
            End If
            
            If (Greatest_Percent_Decrease > file.Range("K" & i).Value) Then
                Greatest_Percent_Decrease = file.Range("K" & i).Value
                Stock_Greatest_Percent_Decrease = file.Range("I" & i).Value
            End If
            
            If (Greatest_Total_Volume < file.Range("L" & i).Value) Then
                Greatest_Total_Volume = file.Range("L" & i).Value
                Stock_Greatest_Total_Volume = file.Range("I" & i).Value
            End If
        
        Next i
        
        ' Populate the Summary Table 2 from the results stored in above variables
        file.Range("O2").Value = Stock_Greatest_Percent_Increase
        file.Range("P2").Value = Greatest_Percent_Increase
        file.Range("O3").Value = Stock_Greatest_Percent_Decrease
        file.Range("P3").Value = Greatest_Percent_Decrease
        file.Range("O4").Value = Stock_Greatest_Total_Volume
        file.Range("P4").Value = Greatest_Total_Volume

        ' Format the Cells
        file.Range("P2").Value = FormatPercent(Greatest_Percent_Increase, 2)
        file.Range("P3").Value = FormatPercent(Greatest_Percent_Decrease, 2)

    Next file


End Sub
Sub Reset()

    For Each file In Worksheets
    
    Dim iCntr
        For iCntr = 1 To 9
            file.Columns(9).EntireColumn.Delete
        Next
    
    Next file

End Sub
