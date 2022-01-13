# stock-analysis
 
## Overview of Project
Our friend Steve is researching green energy stocks for his Parents and was looking for assistance with VBA to analyze stock trends.  His parents have already invested in "DQ", but he wanted to look into other green energy options to diversify their investment.  On initial analysis, DQ showed a dive in stock value, so the research is needed to find a positive investment.  Information was already contained in excel for 2017 and 2018 data, but a script written in Visual Basic could help analyze stock on multiple companies over time with less manual effort.   Data was to be formatted for visual appeal, with buttons to activate the macros on demand.  

## Results
As we initially saw, "DQ" may have been a great choice upon his parents initial investment, however it has taken a turn for the worse in 2018 and drastically lost money.  Using VBA, we were able to analyze the start and ending values of the stocks for each year to find average return on investment.   2017 seemed to be a good year for most green energy companies, however 2018 had an overall negative impact on growth.   Only two companies managed to stay above water; tickers "ENPH" and "RUN".  As seen in the results below, ticker "ENPH" retained substantial growth over both years, and "RUN" was the only company showing an increase in value from 2017 to 2018. Splitting investment among these two companies may fair better for Steve's parents investment.  The VBA script was written for hand input of sheet names based on years, so as future data is available, the analysis can be run fairly quickly.   

!(https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/stock%20outcomes.png)

The initial VBA code, however, was a little redundant and could be improved upon.  Therefore, the time was taken to refactor the code to reduce processing resources and time.  By doing so, we successfully sped up the processing time by 80%.

## Summary
### Why Refactor?


### Refactoring this VBA


Original Run Times

!(https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2017_orig.PNG)
!(https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2018_orig.PNG)

Refactored Run Times

!(https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2017.PNG)
!(https://github.com/catsdata/stock-analysis/blob/main/VBA_Challenge/Resources/VBA_Challenge_2018.PNG)
