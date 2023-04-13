Attribute VB_Name = "Module1"
Sub Datatesting():

'setting my variable as Worksheet
Dim CurrentWs As Worksheet

'creating variables for all the different values and also creating loop for each worksheet.
For Each CurrentWs In Worksheets

Dim totalvolume As Double
Dim ticker As String
Dim openingyear As Double
Dim closingyear As Double
Dim yearchange As Double
Dim yearpercent As Double

'creating the cell where the results for the forloop will show up for
Dim SummaryRow As Long
SummaryRow = 2

'creating a counter for tracking the times the ticker stays the same amd for the opening and closing price and yearly change

totalvolume = 0
openingyear = 0
closingyear = 0
yearchange = 0
yearpercent = 0


'creating a value to count through to the end of all active rows

Dim LR As Long

        LR = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
        
'created a message to make sure that the loop is working for each sheet
        MsgBox LR
        
        'creating headers for the rows using range
       CurrentWs.Range("K1") = "Ticker"
        CurrentWs.Range("L1") = "Yearly Change"
        CurrentWs.Range("M1") = "Percent Change"
        CurrentWs.Range("N1") = "Total Volume Change"
        

                'tracking the opening value of the year
                openingyear = CurrentWs.Cells(2, 3).Value
    'Creating a value to go through the second cell to the last active cell in loop.
            Dim i As Long
            For i = 2 To LR
                
        
                'tracking the changes in row 1 if the current cell and the following cell do not match
                        If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                        
                        'ticker should show up as the cell value
                        ticker = CurrentWs.Cells(i, 1).Value
                        'closing year shows up as the cell value
                        closingyear = CurrentWs.Cells(i, 6).Value
                        
                         yearchange = (closingyear - openingyear)
                      
                                
                                If openingyear <> 0 Then
                                yearpercent = (yearchange / openingyear) * 100
                                
                                'needed to be inside of my if for some reason my information was not resetting
                                
                                End If
                                
    totalvolume = totalvolume + CurrentWs.Cells(i, 7).Value
                                
                        'also condtional that subtracts F:F and C:C when it stops counting same letters
                'need to show ticker changes in a row
                      
                                    CurrentWs.Cells(SummaryRow, 11) = ticker
                                             ' Cells(SummaryRow, 11) = tickercounter dont need to track the lines
                                    CurrentWs.Cells(SummaryRow, 12).Value = yearchange
                                    'converts the numbers to a percent
                                    CurrentWs.Cells(SummaryRow, 13).Value = (CStr(yearpercent) & "%")
                                    
                                    'sends the total sum to 14
                                     CurrentWs.Cells(SummaryRow, 14).Value = totalvolume
                                     
                                     'resets the counter and tracks the changes per the year
                                     ticker = " "
                                     tickercounter = 0
                                     yearchange = 0
                                     totalvolume = 0
                                     yearpercent = 0
                                     closingyyear = 0
                                     SummaryRow = SummaryRow + 1
                                     
                                     'moves down to opening value
                                     openingyear = CurrentWs.Cells(i + 1, 3).Value
                
                                    
                                    
                         
                End If
                
        
                Next i
                
    
    
               
                      'created a conditional loop for colors depending on if the number is positive or negative
                           Dim pr As Long
                        
                           pr = CurrentWs.Cells(Rows.Count, 2).End(xlUp).Row
                           
                            For i = 2 To pr
                            
                            'if its positive
                             If CurrentWs.Cells(i, 12) >= 0 Then
                             CurrentWs.Cells(i, 12).Interior.ColorIndex = 10
                         'its its negative
                                 ElseIf CurrentWs.Cells(i, 12).Value <= 0 Then
                               CurrentWs.Cells(i, 12).Interior.ColorIndex = 3
                        
                        
                             End If
                             
                              Next i
                              
                        
    'Looking for the greatest percent increase

                        CurrentWs.Cells(2, 17) = Application.WorksheetFunction.Max(CurrentWs.Range("M:M"))
        
    'looking for the greatest percent decrease

                        CurrentWs.Cells(3, 17) = Application.WorksheetFunction.Min(CurrentWs.Range("M:M"))
                        
                        
    'greatest volume increase
                        CurrentWs.Cells(4, 17) = Application.WorksheetFunction.Max(CurrentWs.Range("N:N"))
                        
                       'Giving title to the bonus columns
                        CurrentWs.Range("P2") = "Greatest % Increase"
                        CurrentWs.Range("P3") = "Greatest % Decrease"
                        CurrentWs.Range("P4") = "Greatest Total Volume"
                        
                        
                'This for loop goes through each active page and autofits all active columns, it disregards empty
    Dim x As Integer
    For x = 1 To CurrentWs.UsedRange.Columns.Count
                CurrentWs.Columns(x).EntireColumn.AutoFit
    
    Next x
    
     
     Next CurrentWs
     
     
     
     
                        
End Sub

