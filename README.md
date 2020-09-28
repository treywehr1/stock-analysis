# Green Stocks Analysis

## Project Overview
Data was analyzed in Excel using VBA code in order to develop and implement an automated macro that will evaluate the total daily volume, and percentage return of 12 different stocks over the years 2017 and 2018. This macro can be used for any year of comprehensive data collected for these twelve stock tickers. The purpose for this macro is to allow Steve to quickly pull critical information as he researches stock options. Refactoring of a former macro was conducted in order to further optimize the speed, efficiency, and flexibility of the macro. The body of this analysis will address some of the advantages and disadvantages of the refactoring process.
    
## Results of Refactoring and Analysis

While refactoring the stock analysis, some advantages and disadvantages became apparent in the refactoring process. Some were more general conclusions, while others pertained specifically to Steve's macro in VBA or overlapped and affirmed the general pro's/con's.
    
### General Advantages

   - Refactoring code can organize your code and make it easier to read if exectued optimally.
      
   - Refacotring can make your program run more quickly/efficiently.
    
   - Refactoring can keep your code up to date should updates to the software make your original code obsolete.
    
   - Refactoring code offers practice and development to whoever is programming said code
    
### General Disadvantages

   - The larger your script of code gets, the riskier it might be to make large-scale changes to the syntax, as it could be hard to track where discrepencies lie.
    
   - A former programmer could come back into the code and no longer understand the layout of the code he/she had previously worked on.
    
### Green Stocks Specific Advantages

   - After refactoring the code, the program now runs much faster. With the original code script, the time was significantly slower than the refactored code. This is shown by the attached images below. Setting the "tickerIndex = 0" and indexing the tickers instead of running a 'for' loop with an extra variable to run through the tickers eliminated the processing of another loop and, I believe helped with speed significantly.
    
   - Eliminating the ticker 'for' loop also simplified the readability of the code script, since there was only "i" to follow instead of "i" and "j".  
    
                    '4) Loop through tickers            '5) loop through rows in the data
                          For i = 0 To 11                       For j = 2 To RowCount
                          
                                                 vs.
                                                 
                                    
                For tickerIndex = 0 To 11                 For i = 2 To RowCount
                    tickerVolumes(tickerIndex) = 0             If Cells(i, 1).Value = tickers(tickerIndex) Then
                       
### Green Stocks-Specific Disadvantages

   - Refactoring Steve's code took a lot of time on the front end, although it will hopefully help to save time in the long run.
   
   - Refactoring steves code makes this macro less static and easier to apply to further data. This should be a good thing, but perhaps Steve would want to keep      his macro static as stock evaluations can be a sensitive and private business matter.   

## 2017 Code: Original vs Refactored

![Imgur](https://imgur.com/tsGxXHi.png)

![Imgur](https://imgur.com/LZ54VjY.png)

## 2018 Code: Original vs Refactored    

![Imgur](https://imgur.com/sdQj6dl.png)
    
![Imgur](https://imgur.com/0NiVAUk.png)   
   
## Summary of Analysis

In summary, the list of advantages for refactoring seem to heavily out-weigh the disadvantages of refactoring. It was genuinely hard to conceive of disadvantages specific to Steve's VBA code, however there certainly are disadvantages to refactoring in a more general sense. Steve's spreadsheet now processes it's code much more efficiently and is well automated to suit his needs. Most importantly, with cleaner and more concise code, with comments throughout, refactoring has made the actual syntax of this code easier to follow for clearer understanding and more timely ammendments down the line. 
