# Stock-analysis
Analysis of Green Stocks

# Overview of Project
The purpose and background of this project was to help our friend Steve with identifying suitable investments for his parents. His parents want to invest in green energy companies as they believe the sector will offer attractive returns as the world shifts to renewable energy. Steve's parents are particularly interested in DAQO Energy, but they apparently picked it based on their affinity for its stock ticker (DQ). Steve, as a newly minted finance grad, wants to ensure he is responsibly providing financial advice to his parents by doing his research on several publicly traded green energy companies before making a recommendation to his parents. Steve has asked us for help in analyzing past performance of 12 green energy companies.

# Results 
The results show there is a tremendous amount of volatility within the green energy sector. The data from 2017 and 2018 show wide swings in asset prices. The average return in 2017 was 67%, but in 2018 the average return was 8.5%. Additionally, there were wide swings within the index as well; the standard deviation for 2017 was 70% and in 2018 it was 45%. This tells us there is a lot of variation between stocks. A good example of this is Steve's parents’ favorite stock (DQ) which experienced a 198% increase in 2017 before a 62% fall in 2018. There are two stocks that showed strong returns in both years, RUN and ENPH. Both stocks were among the most actively traded during both years, through RUN's returns of 5% in 2017 was well below the index's average. In summary, I would suggest Steve do more research on RUN and ENPH to gain a better understanding as to why these stocks have outperformed their peers over this two-year stretch. I would also want to examine their performance since 2018 to see if the trend continued.

![2017 Analysis](https://user-images.githubusercontent.com/100163289/159179432-37db8823-0ef2-4c72-8403-20127b22a28d.png)
![2018 Analysis](https://user-images.githubusercontent.com/100163289/159179439-8cefd9bb-263d-467e-af6e-8ef7cd56ac65.png)

Below is the refactored code. You can see that it combines several separate functions into a single, streamlined VBA script that is easier to read, follow, and troubleshoot.

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

        yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        TickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
         tickerVolumes(i) = 0
         tickerStartingPrices(i) = 0
         tickerEndingPrices(i) = 0
         
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(TickerIndex) = tickerVolumes(TickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(TickerIndex) And Cells(i - 1, 1).Value <> tickers(TickerIndex) Then
            tickerStartingPrices(TickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
         If Cells(i, 1).Value = tickers(TickerIndex) And Cells(i + 1, 1).Value <> tickers(TickerIndex) Then
            tickerEndingPrices(TickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(TickerIndex) And Cells(i + 1, 1).Value <> tickers(TickerIndex) Then
                TickerIndex = TickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

![image](https://user-images.githubusercontent.com/100163289/159182903-83737844-85c1-467c-9af8-d15ff8444d93.png)

# Summary
The largest advantage of refactoring the VBA code is it combined several functions that had been included in earlier code and combined them into one script. This has several advantages. The refactored code is more efficient, this reducing the run time from nearly a second to around 1/10th of a second. Next, the code is easier to read and understanding, which makes it easier for yourself or someone else to troubleshoot or repurpose the code. Lastly, the code is easier to write to begin with since it is more streamlined which will reduce the number of potential mistakes. 

The disadvantages of refactoring the code are you are redoing work that you have already done. Additionally, refactoring the VBA code could potentially make it more difficult to repurpose the code for other projects that might not need as many functions as utilized in the refactored code. Because it is streamlined, yourself or follow-on programmers might not be able to repurpose the code for other projects as it is fairly bespoke.

The advantages of the original code is each function was a stand-alone code that could easily be troubleshot or repurposed for other projects. It could also be easier to write as no one function was all too terribly complicated. 

The disadvantages of the original code is it took much longer to write as each step was stand alone. This also increased the amount of time it took to run the code, almost a full second vs the 1/10th of a second for the refactored code. While those 9/10th of a second are negligible in this project, a much larger data set might take much longer with the non-refactored code.
