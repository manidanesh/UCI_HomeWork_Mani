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


Dim Max As Double
Dim Min As Double


'Calculate the total number of record in this data set, number of rows which has value'
  last_record = Cells(Rows.Count, 1).End(xlUp).Row


'this counter will be used as a index for the distinct value and add the vlaue in column I, Count 100 means we have 100 unique Tickers in our data set'
counter = 2

Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percentage Change"
Range("l1").Value = "Total Stock Volume"


Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decreased"
Range("o4").Value = "Greatest Total Volume"
Range("p1").Value = "Ticker"
Range("q1").Value = "Value"

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
   
    For J = 2 To last_record
        
        If (Cells(J, 1).Value = Cells(i, 9) And Cells(J, 2).Value = "20160101") Then
          
           year_begin = Cells(J, 3).Value
 
        ElseIf (Cells(J, 1).Value = Cells(i, 9) And Cells(J, 2).Value = "20161230") Then
          
            year_end = Cells(J, 6)

        End If

    Next J
    
 'Update the cell value with the year end - year begin'
    Cells(counter2, 10).Value = year_end - year_begin
 
 'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.'
    If (year_begin <> 0) Then
    
        Cells(counter2, 11).Value = 1 - (year_end / year_begin)
    
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

'The total stock volume of the stock.'
For i = 2 To counter

  For J = 2 To last_record

        If (Cells(i, 9) = Cells(J, 1)) Then

        Total_volume = Total_volume + Cells(J, 7).Value

        End If

  Next J
  
  Cells(i, 12).Value = Total_volume
  Total_volume = 0

Next i


Max = Range("J2").Value
Min = Range("J2").Value

Max_Total_Volume = Range("L2").Value


For i = 3 To 400
    
    If Cells(i, 10).Value > Max Then
    
        Max = Cells(i, 10).Value
        Range("p2").Value = Cells(i, 9)
    
    ElseIf Cells(i, 10).Value < Min Then
    
        Min = Cells(i, 10).Value
        Range("p3").Value = Cells(i, 9)

    End If
    
    If Cells(i, 12).Value > Max_Total_Volume Then
        
        Max_Total_Volume = Cells(i, 12).Value
        Range("p4").Value = Cells(i, 9)
    
    End If
    
    If Cells(i, 10) < 0 Then
        Cells(i, 11) = 0 - Cells(i, 11)
    End If


Next i

Range("q2").Value = Max
Range("q3").Value = Min
Range("q4").Value = Max_Total_Volume

'Change the format of Percenrage change and %increase / desrease '

Range("K:K").NumberFormat = "0%"
Range("q2:q3").NumberFormat = "0%"


End Sub


