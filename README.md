# stock-analysis

## Overview:
The purpose of this analysis was threefold: 
1. To provide a project-based learning experience in developing in VBA.
2. To gain insight into how refactoring of code can improve the quality of the final product in several ways:
   -  **readability**: The ease with which other programmers can understand the code and its organization
   -  **efficiency**: Modifying the program flow to reduce or eliminate unnecessary or superfluous processing
   -  **reuse**: The benefits of reusing "code snippets" whan they have proven to be useful and "bug-free" in previous implementations. 
3. To demonstate by example one desired outcome of refactoring code - improving the speed of execution 

## Results: 

### Comparison of stock performance - 2017 and 2018

Even at a glance, it is clear that the market was kinder to green stocks in 2017 than in 2018. Only two stock (ENPH and RUN) had positive movement in 2018, while in 2017 all but one stock did so. It seems likely that influences on the  market as a whole heavily influenced this patter. Subsequent analysis should compare these stocks with other in defferent sectors to help peel apart the layers of the attributes that affect stock movement. 

![2017 stock performance](https://github.com/DrBuddyO1/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

The following code represent the final struggle with establishing proper execution of the code. The line of = signs was used to establish a visual marker to the site in question and alternative efforts were "commented out" in turn to establish the behavior exhibited by each. 

`            'EDIT OUT======================================================================================
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickersIndex) And Cells(i, 1).Value = tickers(tickersIndex) Then
                tickerStartingPrices(tickersIndex) = Cells(i, 6).Value
            End If
`
### Comparison of execution times - original  / refactored script
Execution times were significantly improved by the refactoring of the code. Times went from ~1.2 seonds to less than a second with this simple refactoring. 

![2018 stock performance with execution times](https://github.com/DrBuddyO1/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary: 
### What are the advantages or disadvantages of refactoring code?
- Advantages: 
  1. Refactoring code is an efficient way to leverage "prior works" to allow subsequent authors to focus on **_improving_** the code and not simply reproducing previous work. 
  2. Refactoring code will, if done thoughtfully, provide a collaborative and community view of the approach and suggest patterns and practices that can be leveraged in other projects. 
- Disadvantages
  1. Starting from an existing working code base might discourage readers or consumers from independent consideration of a problem and approach and reduce the chance that disruptive and revolutionary rather than evolutions thought might be brought to bear. 
  2. Refactoring might encourage or promote the "misapplication" of an approach to other problems and result in awkward or misapplied strategies and principles. 
### How do these pros and cons apply to refactoring the original VBA script?
- Pros: 
  1. Multiple individuals can apply efforts to modifying a successful strategy and a large nunmber of approaches can be kicked off quickly. 
  2. Each effort is faciliated by having a fully working strategy and efforts can be expedited by having a clean and simple approach to debugging small changes rather than risk breaking things in a substantive and vexing way. 
- Cons:
  1. There are a number of alternative strategies for attempting to improve the code from a speed perspective. Most involve efforts to avoid cycling through the full list of tickers by esablishing independent start and end lines for each unique ticker. However, when a successful strategy becomes the starting point for additional attempts, there is sometimes a significant bias toward "tweaks" to a current strategy rather than a new and unique approach. 
