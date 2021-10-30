# Stock Analysis

## Overview of Project

### Purpose
- Steve wants to expand the dataset to include the entire stock market over the last few years.
- Although the VBA we did in [green_stocks.xlsm](green_stocks.xlsm) was working fine for a dozen stocks, it may not work as well for thousands of stocks, and it may take a long time to execute.
- Here I will refactor the code to loop through all the data one time in order to collect the same information that we did before.
- Hopefully, the revised code [VBA_Challenge.xlsm](VBA_Challenge.xlsm) will take shorter execution time.

## Results
![Stock Performance for 2017](Resources/Stock_Performance_2017.png)
![Stock Performance for 2018](Resources/Stock_Performance_2018.png)
- Both the original and the refactored code generated the same stock performance results for 2017 and 2018.
- In 2017, SPWR had the highest 'Total Daily Voume', and DQ had the highest annual return.
- In 2018, ENPH had both the highest 'Total Daily Voume' and the highest annual return. 

![Runtime for refactored 2017](Resources/VBA_Challenge_2017.png)
![Runtime for refactored 2018](Resources/VBA_Challenge_2018.png)
- In terms of the execution time, the refactored code took 0.125s and 0.1132813s for running 2017 and 2018 data respectively.

![Runtime for original 2017](Resources/All_Stock_Analysis_2017.png)
![Runtime for original 20018](Resources/All_Stock_Analysis_2018.png)

- For the original code, it took 0.65625s and 0.6601563s for running the same sets of data.
- The refactored code obviously took shorter time to generate the results based on the same dataset.

## Summary

### Analysis of Outcomes Based on Launch Date
![Outcomes Based on Launch Date](resources/Theater_Outcomes_vs_Launch.png)
- We can see that the months of May and June both have a greater success rate.
- The months of May, Jun, Jul, Aug and Oct have similarly high number of failure rate.
- The canceled rate in January is the highest.
- Here is the link to [Kickstarter_Challenge.xlsx](Kickstarter_Challenge.xlsx).

### Analysis of Outcomes Based on Goals
![Outcomes Based on Goals](resources/Outcomes_vs_Goals.png)
- We can see the higher percentage of success for Goals in less than 1k, 1k-1.5k, 35k-40k, and 40k-45k (75.81%, 72.66%, 66.67%, 66.67%).  
- Higher percentage of failed for the Goals in 25k-30k, 45k-50k =, and over 50K (80%, 72.73%, 100%, 87.5%). 
- Here is the link to [Kickstarter_Challenge.xlsx](Kickstarter_Challenge.xlsx).

### Challenges and Difficulties Encountered
- One challenge is to setup `COUNTIFS` formulas in the sheet "Outcomes Based on Goals". By copy and paste the formulas to other columns without using absolute cell reference will require lots of updates in the formula. By adding absolute cell reference (using the shortcut **Fn-F4**), it makes my life easier!
``` 
=COUNTIFS(Kickstarter!D:D,">=5000",Kickstarter!D:D,"<=9999",Kickstarter!F:F,"successful",Kickstarter!R:R,"plays")
=COUNTIFS(Kickstarter!$D:$D,">=5000",Kickstarter!$D:$D,"<=9999",Kickstarter!$F:$F,"successful",Kickstarter!$R:$R,"plays")
```
- Another challenge is to reduce the Excel file size. The Excel file size is over 45M, and it is mainly due to too large of "used range" in the worksheet "Successful US Kickstarters" and "Failed US Kickstarters". I google the solution and find this https://www.excelefficiency.com/reduce-excel-file-size/#1_Remove_8220blank_space8221_in_your_sheets to describe how to clear the unneccessary "used range". Afterwards, the file size now reduce to under 3M!

