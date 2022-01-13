# stock-analysis
 
## Overview of Project
Our friend Steve is researching green energy stocks for his Parents and was looking for assistance with VBA to analyze stock trends.  His parents have already invested in "DQ", but he wanted to look into other green energy options to diversify their investment.  On initial analysis, DQ showed a dive in stock value, so the research is needed to find a positive investment.  Information was already contained in excel for 2017 and 2018 data, but a script written in Visual Basic could help analyze stock on multiple companies over time with less manual effort.   Data was to be formatted for visual appeal, with buttons to activate the macros on demand.  

## Results
As we initially saw, "DQ" may have been a great choice upon his parents initial investment, however it has taken a turn for the worse in 2018 and drastically lost money.  Using VBA, we were able to analyze the start and ending values of the stocks for each year to find average return on investment.   2017 seemed to be a good year for most green energy companies, however 2018 had an overall negative impact on growth.   Only two companies managed to stay above water; tickers "ENPH" and "RUN".  As seen in the results below, ticker "ENPH" retained substantial growth over both years, and "RUN" was the only company showing an increase in value from 2017 to 2018. Splitting investment among these two companies may fair better for Steve's parents investment.  The VBA script was written for hand input of sheet names based on years, so as future data is available, the analysis can be run fairly quickly.   

![Stock Results](https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/stock%20outcomes.png)

The workbook can be accessed [HERE](https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/VBA_Challenge.xlsm)

The initial VBA code, however, was a little redundant and could be improved upon.  Therefore, the time was taken to refactor the code to reduce processing resources and time.  By doing so, we successfully sped up the processing time by 80%.   Nested loops were improved upon.  The 2017 data run times went from 1.078 seconds to 0.172 seconds, and the 2018 data run times dropped from 1.035 seconds to 0.215 seconds.  Imagine how much time and resources that will save for future and/or expanded datasets!

Original Run Times

![2017](https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2017_orig.PNG)
![2018](https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2018_orig.PNG)

Refactored Run Times

![2017](https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2017.PNG)
![2018](https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2018.PNG)


## Summary
### Why Refactor?

Refactoring code, of any type, can improve overall efficiencies.  Just as writing the VBA code initially was to improve processing time on analyzing the stock data manually in Excel.  By making it more efficient, and ensuring comments are added to explain how the code it working, its easier to reuse in the future and make more accessible to other users.  Refactoring will also reduce the human error aspect provided datasets are added in the same format and the least amount of interaction is required.  Initial up front time investment ultimately reduces the overall cost of labor and effort in the long run.  And good code can be shared and reused or tweaked for other projects as well.

However, refactoring isn't always necessary.  If the project is a one time need, there's no sense refactoring if it works, and won't be needed again.  The effort should always be weighed against the need.

### Refactoring this VBA

Code for this project was refactored to reduce run times by nesting loops where able.  As seen above, run times substantially improved.  Although the initial code worked for the small set of stocks provided, it may not have been the best of code should we need to analyze stocks outside of the small green energy genre.  

Refactoring this code took over 10 hours even with intitial guidance provided.  I would say this would have been a con if the process of debugging and code research didn't improve upon my understanding of how the code works and where I was having issues.  So in the long run, this not only improved run times for Steve and his parents, it made me a better programmer.  
