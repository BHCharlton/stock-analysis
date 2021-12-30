# Stock Analysis

## Overview of Project
Our client, Steve, wanted to continue using the workbook we had made for him, but needed to expand his research to the entire stock market as a way to determine what stocks were worth investing in.  Thus, it was important to ensure that the VBA code would still be able to run quickly and efficiently when processing much larger sets of data.  The purpose of this challenge was to "refactor" the code that had already been written throughout the Module to process stock information for an entire year that would provide the total daily volume for the selected year as well as the overall return of each stock for that year.

## Process
Market information for twelve different stocks during the years 2017 and 2018 is used in this challenge to test the efficacy of the refactored VBA code.  The data provides a ticker for each stock along with the date it was issued, an opening and closing price for that stock, the all-time high and low price, as well as the adjusted closing price and total volume.  I began by downloading the provided "starter code," then made the appropriate additions to the code, as outlined in the challenge, to refactor the script to run more efficiently.  I then used the provided market information as a means for testing the efficacy of my additions.

## Results
The following images show the results of running the refactored VBA code for the years 2017 and 2018.  These images are also available in the repository:
### 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/93561592/147627572-0c7d0703-aae8-43fc-8ff8-a17c0ce2ce7c.png)
### 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/93561592/147627602-fb90d7d5-dd26-4cca-a7da-14e3c6b0431a.png)


## Summary
The images above show that the refactored VBA code provided the desired results in less than one second, which is nearly four times faster than the processing time before refactoring the code. Obviously, the efficiency obtained through refactoring the code is the most significant benefit in this case.  In addition, refactoring the code helps to provide better overall organization and layout of the entire script.  However, refactoring code can also open the door to creating bugs that our tests don't catch which could lead to critical mistakes when assessing the results.  Fortunately, for this challenge, we had pictures and on-line guidance to compare our result with and insure that it was running correctly. 
