# Challenge 2 , Deliverable 2 - VBA of Wall Street

## Overview of the Project

### Purpose

I am aiming to improve the efficiency of my subroutine "AllStocksAnalysis" by refactoring my code. Efficiency is measured by how long the subroutine takes to finish. The code is ran, and at the end there is a text box telling me how long it took to complete running. 

### Background

The subroutine "AllStocksAnalysis" originates from work done in the background, where I run an initial analysis on a stock "Daqo" with the ticker "DQ". I wanted to find the daily volume and the yearly return of this stock with the information that was provided in order to help his parents who had invested in this company, to measure the performance of the stock over a period of two years. Over time, the subroutine evolved into also measuring the performance of the rest of the stocks that were included in the data provided over a period of one year of the user's choosing. By the end of the module, I was able to measure the performance of my subroutine by the code's runtime duration. Here, I'm comparing the duration of the code's runtime before and after the refactoring of the code.

## Results
The results of the refactor of the subroutine are astounding! I reduced the runtime of the code by a factor of 10. You can see in the images below what the times to run the code for the years 2017 and 2018 where before the refactor as well as times to run the code after the refactor.

Before:

![Before refactoring: 2017](https://github.com/etdirksen/stock-analysis/blob/main/Resources/Before_2017.png) ![After refactoring: 2018](https://github.com/etdirksen/stock-analysis/blob/main/Resources/Before_2018.png)

After[^1]:

![After refactoring: 2017](https://github.com/etdirksen/stock-analysis/blob/main/Resources/After_2017.png) ![After refactoring: 2018](https://github.com/etdirksen/stock-analysis/blob/main/Resources/After_2018.png)


### Why the new subroutine is faster

The improvement comes mainly from adding a buffer to the code. Before the refactor, our code would switch between two different worksheets every time we output any data. This is what caused it to be so slow. In the refactor, I added a buffer; a buffer is just a thing that temporarily holds data before unloading it somewhere else. By utilizing a buffer, I can drastically cut down how many times the initial code needs to switch between worksheets in order to update information.

The __old code__ looked like this:
```
'4) Loop through tickers
For i = 0 To 11
    ticker = tickerArray(i)
    totalVolume = 0

    '5) loop through rows in the data
    Worksheets(yearValue).Activate
    For r = rowStart To rowEnd
    
        '5a) Get total volume for current ticker
        If Cells(r, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(r, 8).Value
        End If
    
        '5b) get starting price for current ticker
        If Cells(r - 1, 1).Value <> ticker And Cells(r, 1).Value = ticker Then
            startPrice = Cells(r, 6).Value
        End If
        
        '5c) get ending price for current ticker
        If Cells(r + 1, 1).Value <> ticker And Cells(r, 1).Value = ticker Then
            endPrice = Cells(r, 6).Value
        End If
        
    Next r
        
    '6) Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endPrice / startPrice) - 1

Next i

```

So what is happening here? The program is looping through every single row in the data table that we have and each time a piece of information for the current ticker is updated, like say the total volume for the current ticker, then the variables in our code are updated. When it has reached the end of data that is relevant for our current ticker, the code will switch to another worksheet and output that information. It will then switch back to the worksheet that contains the data table, reset the variables, continue looping through the data, and will start updating the variables based on the next ticker.

This is slow because the program is writing data to a new worksheet for every ticker that exists in the data table. If we had 100 tickers instead of about 10, this code would take 10 times longer to run, or about 7.5 seconds per year that we selected. With how fast modern computers are today, this is extremely slow. To remedy this, I need to reduce the amount of times that I need to write to a new sheet.

The __new code__ looks like this:
```
''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If (Cells(i - 1, 1).Value <> Cells(i, 1).Value) And (Cells(i, 1).Value = tickers(tickerIndex)) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        If (Cells(i + 1, 1).Value <> tickers(tickerIndex)) And (Cells(i, 1).Value = tickers(tickerIndex)) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
```

The major difference is that I implemented a way to output the data to a new worksheet, _outside of the for loop_. So even if there were 1,000 tickers that we were gathering data for, the code will only switch to the other worksheet to output all of this data only once.

## Summary

### Advantages of refactoring in general
There are many advantages to refactoring code in general. One advantage that has been displayed here is reducing the runtime of the code. With better logic, code can be written to run orders of magnitude faster than before. This can also mean that the code can maximize the use of the available resources. Similarly, refactored code can use less available resources while maintaining the same or a similar runtime. Another advantage of refactoring code is improving readability and editability. Good refactored code would reduce how many things you need to change in order to fit your use case, how long it takes to find and change them, or both.

### Disadvantages of refactoring in general
Refactoring code is a time consuming and often difficult task. Reducing the time that it takes for any code to run requires forming a better solution to the problem at hand. The greater the solution, the greater the gains when the solution is implemented. Oftentimes, forming and implementing that better solution takes some order of magnitude of time and many different resources. In rare cases, there may not be a better solution to use as there are hard limits to the simplicity of the code. As code gets simpler and simpler, readability also decreases. If the code is extremely complex but offers the best solution, readability and editability are decreased.

### Advantages of the original & the refactored VBA script
The advantages of the old code are that it stood as the foundation for the new code to be based on. There was not much changed in the new code in order to reduce the runtime. Of course, there are no advantages to the old code when compared to the new code. 

The advantages of this refactor are obvious - to summarize, the new code will almost always take about 1/10 of the time to finish running compared to the old code because of better logic (in our case, this is how many times we output the data). 

### Disadvantages of the original & the refactored VBA script
The disadvantages of the old code are that it was slow. If you imagine having thousands of different stocks to analyze, this code would be running for several minutes per year.

I think the disadvantages of the refactored code are that while it is much faster and more efficient than the old code, it is only useful for the 12 stocks that are currently listed. You would have to change the size of each array and possibly the number of arrays to fit more stocks if there were more, and you would have to write out each stock ticker by hand as every stock ticker is hard coded. It would be much better to create variables to store the different stock tickers as they are found in the data table. Of course the problem with implementing this better logic is the time that it would take to write a solution that fits within the rules of the language.

Also, the data table for the old and the refactored code has sorted all of the data for each stock ticker. When working with various different data sets, this may not always be the case.

[^1]: In standard notation, this is: 0.06982422 seconds for 2017 and 0.07397461 seconds for 2018.
