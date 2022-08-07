# Kickstarter-analysis
Performing analysis on Kickstarter data to uncover trends
VBA – Written Analysis
Overview of Project
Background:
What has already been prepared for the DQ analysis was insightful and helpful to Steve regarding whether or not to invest in that particular stock. While investing in DQ is not in his best interest, perhaps there is stock that would be a beneficial investment. Steve now wants see data for the entire stock market over the last few years. Because there is usable code already, we now only need to edit or refactor the code to suit Steve’s needs for additional, insight and information into the stock market for a specified period. 

Purpose:
The purpose of the refactored DQ code is to expand the current dataset to provide Steve with a view of the total volume of investments and percentage of loss and/or gain of all stocks over 2017 and 2018. Not only will we provide a broader scope of information to suit Steve’s new interest; that information will be presented in a more efficient format. 

Results
Steve is interested in looking at the volume and return of the entire stock market for the years: 2017 and 2018. Based on the dataset for both years, there are two attractive stocks worth investing in: RUN, yielding a return of 84% and ENPH, with a return of 81.9%. Based on the findings in our table ENPH also has the highest volume, proving to be a popular stock. Interestingly, SPWR has the second highest total daily volume of 538,024,300 but one of the worst performing amongst the stock. 
					
This dataset was achieved by refactoring the existing code for Steve’s previous DQ Analysis. The following output arrays were used to establish our variables for the stocks starting price, ending price and volume: 
  Dim tickerVolumes(12) As Long
 Dim tickerStartingPrices(12) As Single
 Dim tickerEndingPrices(12) As Single

The additional code below was used to conveniently loop through all the rows in our spreadsheet. This block of code increases convenience by including several if/then statements, adding to the precision of our results:
       


Code Preview

 For rowstart = 2 To RowCount
           If Cells(rowstart, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(rowstart, 8).Value
            
            End If
      Based on the results of our dataset achieved with the aforementioned code; the most lucrative stock for Steve and/or his parents to invest in would be, RUN. 
        
Summary
Advantages & Disadvantages of Refactoring Code
One of the greatest and most obvious benefits of refactoring code is the that it takes less time to do so. Additionally, refactoring helps in finding mistakes, bugs and “cleans up” what may now look like messy data. Refactoring can also make the data easier to understand. 
Time spent can also be a disadvantage when refactoring code. It can take a lot of time to make sense of existing and sometimes old code. It is also possible to introduce new bugs to existing code, defeating the purpose of refactoring to create greater efficiency. 
Advantages & Disadvantages of the Original Code
The original code served a narrow purposed and provided a great foundation for the additional code. The disadvantage of the original code is how much time it took to create and the kinks that need to be worked out to make the code run efficiently and provide the desired output. 
Advantages & Disadvantages of Refactored Code
The refactored code was simplistic though broader in the information it provided. It appears more organized and easier to understand than the original code. There were however a few trial-and-error moments in the refactored code regarding syntax and structure before running smoothly. 
 

Footer
© 2022 GitHub, Inc.
Footer navigation
Terms
Privacy
Security
Status
Docs
Contact GitHub
Pricing
API
Training
Blog
About
