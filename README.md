# **Improving Code Performance on Stock Analysis in VBA**

## **Overview of Project**

### **Purpose**
One of the first checkpoints on the road to creating a good program, no matter
how small it might be, is making it work. We ask *can it do the task(s) it was*
*designed to do?* Once complete, work is done to ensure that the program is
maintainable. Next comes the consideration of performance. Here, the question
becomes *how fast is it and how much resources does it use?*. The purpose of
this project concerns the first part of this performance question. Once our
program works and is maintainable, what methods can we use to speed it up? The
analysis below employs refactoring to achieve better execution time with a
discussion on why, in general, it is a good idea in any project to look for
opportunities to refactor code.


### **Context**
The data of interest that will serve as the context for this code performance
analysis are stocks from 2017 and 2018 from 12 different green energy companies
([VBA_challenge.zip](/VBA_challenge.zip)). The objective is to see which stocks
have been producing positive yearly returns with the idea that these would be
chosen for personal investment. Subroutines were created in VBA to automate the
calculations of total daily volume and yearly return. Code execution time is
printed out to the user each time the analysis is run.

## **Analysis and Results**
### **Original Subroutine**
The subroutine `AllStocksAnalysis` was created to automate creation of the
total daily volume and return table for each stock ticker for a given year. The
program starts by taking input from the user as to what year to analyze. Then,
choosing one stock ticker at a time (first loop index `i`), the entire data set
is looped over by row (second loop index `Row`) to gather all the volume and
pricing data that match that ticker. Here is pseudocode showing a
representation of this nested structure:
```
For i = 0 To 11
    ticker = tickers(i)
    ...
    For Row = 2 To RowCount
    ...
    Next Row
Next i
```
Conditional logic used within the inner loop determines the yearly starting and
ending prices by comparing the ticker values of the rows before and after the
current row in the loop. If the ticker values are different, this signals the
beginning or end of the data for a given ticker. Script timing was measured
starting after the user input year was collected until the data was output to
the table. Here are the results tables from 2017 and 2018 produced by this
subroutine:

<p align="center">
  <img src="/resources/2017_all_stocks_analysis_table.svg">
  <img src="/resources/2018_all_stocks_analysis_table.svg">
</p>

Clearly the tickers ENPH and RUN have had strong positive returns in both years
with ENPH being the strongest. Moving on to the performance of this script, the
following message box outputs show the execution time for each year's analysis:

<p align="center">
  <img src="/resources/VBA_challenge_2017_non_refactored.svg">
  <img src="/resources/VBA_challenge_2018_non_refactored.svg">
</p>

### **Inefficiency of the Original Subroutine**
Upon review of the original subroutine, it is observed that the entire table of
stock data is looped over for each of the tickers. Each stock ticker has only a
small number of rows compared to the total number of rows, so most of the loop
is spent evaluating rows that do not contain data relevant to that ticker. This
inefficiency represents an opportunity to refactor the program to avoid useless
row evaluations and reduce the number of times the table is looped over.

### **Refactored Subroutine**
To reduce the number of times the code loops through the table arrays can be
used to collect and store the volume and pricing data at intermediate points in
the loop. These intermediate points occur when the ticker value changes from
one to the next. The subroutine `AllStocksAnalysisRefactored` employs this
strategy and the result is a refactoring of the original subroutine using only
one loop through the table of data. Below is pseudocode of the refactored
structure. Here more detail is shown than in the previous code snippet for
clarity.
```
tickerIndex = 0

Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
...
For Row = 2 To RowCount
    update tickerVolumes(tickerIndex)

    If <at beginning of ticker rows> Then
        update tickerStartingPrices(tickerIndex)
    End If

    If <at end of ticker rows> Then
        update tickerEndingPrices(tickerIndex)
        tickerIndex = tickerIndex + 1
    End If
Next Row
```
The pseudocode shows that once the rows of the current ticker are finished (the
second `If` statement in the `For` loop) `tickerIndex` is incremented, allowing
the arrays to move on to storing the next ticker's data. When each row is
evaluated for its data in the `For` loop, it is always relevant to the current
ticker associated with `tickerIndex`, thus there are no wasted row evaluations.
Here are the timing results for the refactored subroutine for both years:

<p align="center">
  <img src="/resources/VBA_challenge_2017_refactored.svg">
  <img src="/resources/VBA_challenge_2018_refactored.svg">
</p>
The results show that the refactored version produced approximately an 85%
reduction in execution time.

## **Summary**
### **Benefits**
The analysis demonstrates a clear benefit to refactoring the code in this
application. In general, evaluating where a program is doing most of its work
and then stepping through the logic to see if there are any wasted actions can
reveal ideal locations for refactoring. Clear benefits can include reducing
nesting or total lines of code which enhance readability. Also, as in this
application, wasted actions can be reduced or removed thus speeding up the
time to compute. It is not as clear here, but as the size of the data set
grows, it is possible that the inefficiency scales along with it, thus making
the performance gains from refactoring even more impactful.

### **Drawbacks**
It is possible that over emphasis on refactoring can lead to very concise code
that is not very self-descriptive. If the performance benefits are substantial,
then well-commented concise code is likely worth the effort. If, however, there
is minimal gain or no gain at all, sacrificing self-describing code might have
a negative impact on collaboration or on picking back up on the work at a later
date. Another drawback is simply the time it might take to refactor. First, it
takes time to locate where code could be improved and then it has to be written
and documented. The time to complete this might not be trivial depending on the
application.

This project primarily benefitted as the performance gains were significant and
it was relatively easy to locate where there was ineffeciency in the program.

