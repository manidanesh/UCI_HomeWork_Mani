Sub Analytics():


'define a counter to keep track of the tickers index '
Dim counter As Integer
'Define a value to store the start date of a year'
Dim year_begin As Double
'Define a value to store the end date of a year'
Dim year_end As Double

Dim diff_year As Double
'This is a temprory value which will be used for identiy the Distinct Ticker'
Dim current_value As String


'Calculate the total number of record in this data set, number of rows which has value'
  last_record = Cells(Rows.Count, 1).End(xlUp).Row


'this counter will be used as a index for the distinct value and add the vlaue in column I, Count 100 means we have 100 unique Tickers in our data set'
counter = 2



For i = 2 To last_record

        'this condition is checking if a new value exisit in the data set for tickers
        If Cells(i, 1) <> current_value Then
          
            Cells(counter, 9).Value = Cells(i, 1).Value
            current_value = Cells(i, 1).Value
      
      'This value will count the total number of unique Tickers in the data set '
          counter = counter + 1

        End If
        
Next i
'End of identifying '


'Early change from opening price at the beginning of a given year to the closing price at the end of that year.'
  counter2 = 2

For i = 2 To counter
   
    For j = 2 To last_record
        
        If (Cells(j, 1).Value = Cells(i, 9) And Cells(j, 2).Value = "20160101") Then
          
           year_begin = Cells(j, 3).Value
 
        ElseIf (Cells(j, 1).Value = Cells(i, 9) And Cells(j, 2).Value = "20161230") Then
          
            year_end = Cells(j, 6)

        End If

    Next j
    
 'Update the cell value with the year end - year begin'
    Cells(counter2, 10).Value = year_end - year_begin
 
 'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.'
    If (year_begin <> 0) Then
    
        Cells(counter2, 11).Value = year_end / year_begin
    
    End If

  counter2 = counter2 + 1
  year_begin = 0
  year_end = 0

Next i
'End of calculation for Yearly change and %of change for the openning price at beginning vs closing at the end of that year.'

'Conditional formatting that will highlight positive change in green and negative change in red.'
For i = 2 To counter

  If Cells(i, 10).Value > 0 Then

      Cells(i, 10).Interior.ColorIndex = 4

  Else

      Cells(i, 10).Interior.ColorIndex = 3

  End If

Next i
'End of the loop for formatting the column Yearly_Change'



'Start section calculating total volume and update the column 12 '
current_value = Cells(2, 9)
Total_volume = 0


'''''''
'The total stock volume of the stock.'
For i = 2 To counter

  For j = 2 To last_record

        If (Cells(i, 9) = Cells(j, 1)) Then

        Total_volume = Total_volume + Cells(j, 7).Value

        End If

  Next j
  
  Cells(i, 12).Value = Total_volume
  Total_volume = 0

Next i


End Sub





