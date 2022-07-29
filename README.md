# VBA-challange
Sub Alphastock()
Dim tickername As String

Dim tickerboxrow As Integer
tickerboxrow = 1
Dim closeprice As Integer
Dim openprice As Integer
Dim yearlychangeprice As Integer
Dim yearlychangepercent As Integer
 Dim totalvolume As Integer
            totalvolume = 0

'loop thru every entry

    
    For r = 1 To lastrow
'check if we are still within the same ticker

    If Cells(r + 1, 1) <> Cells(r, 1) Then
'set the tickername
    
    tickername = Cells(r, 1)
'close price yearly + open price yearly
    closeprice = Cells(r, 6).Value
    openprice = Cells(r, 3).Value
'yearly change print IT!
    yearlychange = closeprice - openprice
'yearly change percentage (easy peasy)
    yearlychangepercent = yearlychange / openprice
'print the ticker in the summary table
     Cells(tickerboxrow, 9) = tickername
'print the total amount in summary
     Cells(tickerboxrow, 10) = totalvolume
           
     
     Range("j" & tickerboxrow).Value = totalvolume
 'volume total to summery
 
    tickerboxrow = tickerboxrow + 1
 
'percent change to sumery
     Cells(tickerboxrow, 11) = yearlychangepercent
    
    yearlychangepercent = Cells(tickerboxrow - 1, 11)
' reset the open price
openprice = Cells(r + 1, 3)
' reset the volume total
volumetotal = o
Else

'

   
    
Next r
