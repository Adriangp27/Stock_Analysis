# Stock_Analysis
## Overview

An analysis of the Stocks dataset that was gathered over the past two years is requested by the client. The results are displayed through a table that displays the Ticker, the Total Daily Volume, and the Return for each year. By looking at this table, the investor will be able to see what the success rate is for each stock to make an informed decision.  


## Results

In spite of the fact that the original code works as intended, it is always possible to refactor it for a more efficient one. For example, faster runtimes can be achieved when both codes are processed simultaneously. Based on the 2017 stock market, the refactored script runs for about 0.1 seconds or less before and after refactoring (Note that the refactored runtime is running below 0.1 seconds, therefore it is demonstrated in scientific form).

![All_Stocks_Original](https://user-images.githubusercontent.com/99752443/159099053-f34625b2-f5cc-4469-afe7-4878b040ecc2.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/99752443/159099077-4ab84f57-28ef-48b9-b2ad-f70f2008a6f3.png)


## Summary
###### Advantages

In addition to being able to handle more features if the client requests them in the future without negatively impacting the performance on the client, the script is optimized in order to create a more organized and readable code for the developer and others working on it in the future. Rather than accessing all of the tickers array via a variable, ticker, a tickerindex can be used to access it directly without having to set the array's value. As a result, it is now easier to follow the script and the code also occupies less memory since it eliminates one step. Furthermore, since the index can now be accessed and set by the user, more arrays can now be created to hold the value of the totalVolumes, tickerStartingPrices, and tickerEndingPrices. Consequently, the client is left with a placeholder for all the data that they need, even though the same logic can also be applied if they wanted more specific data pulled from the dataset; all of which are accessible via the index created.   
```
 Dim tickerIndex As Integer
    tickerIndex = 0
    
    'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    'Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(i) = 0
    
    Next i
    'Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        'Increase volume for current ticker
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        'Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        'check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
            ElseIf Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
         'Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            End If
         
    Next i
    
```

###### Disadvantages

In some situations, it's best to leave original scripts alone. For instance, refactoring can't fix broken code. The whole point of refactoring is to create a more reliable code base that you can build on. Often, when the original code works, refactoring might introduce new bugs. Going back and refactoring the script will not be possible before the initial deadline. In this way, developers can prevent code rot by continuously adding features and optimizing their code.
