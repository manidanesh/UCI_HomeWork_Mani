Sub Analytics():


'define a counter to keep track of the tickers index '
Dim counter As Integer
'Define a value to store the start date of a year'
Dim year_begin As String
'Define a value to store the end date of a year'
Dim year_end As String

Dim diff_year As Double
'This is a temprory value which will be used for identiy the Distinct Ticker'
Dim current_value As String


'In this condition section, I will check to see if there is any value sets in the new Ticker column'
  If Cells(2, 9).Value = "" Then
    Cells(2, 9).Value = Cells(2, 1).Value
    'Counter_value, is the latest value in the data set which is diffrent from the previus value '
    current_value = Cells(2, 1).Value
    counter_value = 3
  End If

'Calculate the total number of record in this data set, number of rows which has value'
last_record = Cells(Rows.Count, 1).End(xlUp).Row


'this counter will be used as a index for the distinct value'
counter = 3

For i = 2 To last_record

        If Cells(i, 1) <> current_value Then
            Cells(counter, 9).Value = Cells(i, 1).Value
            current_value = Cells(i, 1).Value
      
      counter = counter + 1
        End If

Next i
'End of identifying '


counter2 = 2

For i = 2 To counter
    For j = 2 To last_record
        
        If (Cells(j, 1).Value = Cells(i, 9) And Cells(j, 2).Value = "20160101") Then
           year_begin = Cells(j, 3).Value
 
        ElseIf (Cells(j, 1).Value = Cells(i, 9) And Cells(j, 2).Value = "20161230") Then
            year_end = Cells(j, 6)

        End If

    Next j
    
  Cells(counter2, 11).Value = year_begin - year_end
  Cells(counter2, 12).Value = Cells(i, 9)
  counter2 = counter2 + 1
  
  year_begin = 0
  year_end = 0
Next i



End Sub




