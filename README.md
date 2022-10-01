# stock-analysis

## Purpose
The purpose of this analysis is to review an entire dataset of stocks and understand their performance. This analysis was done by refactoring code, and improving it so it can analyze different years and more stocks if neccesary.


## Results

### Images

![2017]()

1[2018]()

### Code

   '1a) Create a ticker Index
    
    Dim tickerindex As Integer
    tickerindex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For x = 0 To 11
    
        tickerVolumes(x) = 0
        
   Next x
        
    ''2b) Loop over all the rows in the spreadsheet.
    
        For i = 2 To RowCount
            
        
            If Cells(i, 1).Value = tickers(tickerindex) Then
            
                tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
                
            End If
            
            
            '3a) Increase volume for current ticker
                    ' Solved in 2b
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
                
            If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            
                tickerStartingPrices(tickerindex) = Cells(i, 6).Value
    
            End If
                
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
                
            If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                tickerEndingPrices(tickerindex) = Cells(i, 6).Value
            End If
    
            '3d Increase the tickerIndex.
                
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then            
                tickerindex = tickerindex + 1

            End If
        Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        'Activate the sheet
        Worksheets("All Stocks Analysis").Activate
                              
        'Write results in the active sheet
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

## Summary

- There is a detailed statement on the advantages and disadvantages of refactoring code in general
Refactoring code allows a faster analysis as the user doesn't need to create it from scratch. On the other side, if the user don't understands the code entirely, it could take longer than expected.

- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script
The advantage of refactoring code in VBA is that you already have a visual result in excel, which can help you understand better the code if it is the first time using it. As well as in any language, don't understanding the code entirely could make it a slower process than writting it from scratch.