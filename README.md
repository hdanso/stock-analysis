# Analyzing Stock Trends with VBA

## Overview of Project

### Purpose

Visual Basic for Applications (or VBA for short) is a powerful tool that can be used to automate manual tasks in Excel and perform analyses using complex logic. Initially, stock analysis was performed using macros that automated the output of the total daily volume and yearly return of 12 green stocks. The original script was then refactored to improve its logic. The purpose of this project was to improve the efficiency and logic of the original code, thus decreasing the overall run time of stock analyses.

## Results

### Stock Performance

Based on the analysis of 2017 and 2018 stock using total daily volume and yearly returns as performance metrics, 2017 stocks had a significantly better performance compared to 2018 stocks. 11 out of 12 stocks had an overall positive return in 2017, with the top performers being DQ, with an overall return of 199.4%! The highest total daily volume in 2017 was SPWR, with a total daily volume of $782,187,000.

![VBA_Challenge_2017_Performance](https://user-images.githubusercontent.com/96188669/188319844-97e4101d-699b-43f1-9a64-019a3c23aeb9.png)


There was a stark difference in 2018 stock performance. Comparatively, only 2 out of the 12 stocks had an overall positive return in 2018. The top performer was RUN, with an overall return of 84.0%. In contrast, the lowest performer in 2018 was DQ, with an overall return of -62.6%. The highest total daily volume in 2018 was ENPH, with a total daily volume of $607,473,500. 

![VBA_Challenge_2018_Performance](https://user-images.githubusercontent.com/96188669/188319854-789c61be-cef9-4609-a45d-bc3589d1ef0b.png)


Comparing the 2 metrics undoubtedly shows that 2017 was a better year for green stock performance than 2018.

### Script Execution Times

 The original code had a script run time of 1.96875 seconds for 2017 stock analysis and 1.980469 seconds for 2018 stock analysis.
 
 
![VBA_Challenge_2017_Original_Code](https://user-images.githubusercontent.com/96188669/188319879-9df328b9-a0be-4133-b26e-dd9cf7dbf87d.png)
![VBA_Challenge_2018_Original_Code](https://user-images.githubusercontent.com/96188669/188319884-d080b7b4-5960-4508-aaec-1c5a2efb593c.png)



In comparision, the refactored code had a script run time of 0.3046875 seconds for 2017 stock analysis and 0.34375 seconds for 2018 stock analysis. 


![VBA_Challenge_2017](https://user-images.githubusercontent.com/96188669/188319916-456c4fc0-4514-4570-ba08-0abb42f8d16f.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96188669/188319923-76344b88-070a-49c1-8e1d-f16455c6d198.png)



Comparing script execution times, the refactored script had a much faster run time than the original code used for the analysis.

## Summary

### Advantages of Refactored Code

In general, refactoring code is best practice to improve the overall speed, logic, and structure of your original code. When a code is refactored, less steps and memory is used to perform the same task as the original code. This improves the overall run time, as the system does not have to work through as many steps to produce an output. This is a major advantage when working with larger datasets, as increasing code efficiency will reduce the overall run time. In addition, refactored code is advantageous when sharing codes with other users, as less steps in your code make it easier to read and understand for future use.

### Drawing Efficient Stock Insights with Refactored Code

With regards to the stock analysis, refactoring the original code resulted in a faster run time for both 2017 and 2018 stock analysis, and the completion of the same analyses with added efficiency. Refactoring this code also made it easier to catch and correct errors in the code, as the overall readability increased since the refactored code contains less steps and process repetition. By refactoring this code, future users will be able to re-use it to analyze larger datasets with better efficiency.
