# Stock-Analysis

## Overview of Project
Steve wants the ability to quickly scan the entire stock universe for returns and total volume. He will use this tool to make stock recommendations for his parents. The first VBA we built accomplishes this but, in an inefficient manner. The finished VBA can be used on an entire stock universe very quickly and accurately.

### Purpose
The goal of this project is to build a VBA script and tool that can analyze every stock over many years in an efficient process.

## Analysis
VBA was used to perform analysis. There are two versions of VBA with differences in looping through raw data. We seek to see which VBA script is better. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/103381098/166058249-ac09081a-f3af-4755-8df8-ddf56769f6c8.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/103381098/166058276-8f97f267-1823-4e7f-b31e-6cab42a32826.png)

### Stock Performance
Overall, the basket of stocks analyzed performed better for 2017, only one stock had a negative return. However, 2018 was significantly worse with only two stocks having a positive return. ENPH and RUN were the only companies to have positives returns both years. During 2018, the only two positive stocks also were in the top three for total daily volume. There weren’t any strong correlations for returns and total daily volume for 2017.

### Execution Times
The refactored VBA script vastly outperformed the original script. The new script completed 2017 in 6.640625E-02 and 2018 in 0.0625 seconds. The old script completed 2017 in 0.6679688 and 2018 0.6289063 in seconds. Both years are just over 10 times faster using refactored script. While the absolute time might not seem like much of a difference over 12 stocks it will save significant time when applied to hundreds or thousands of stocks. The main difference with new script is that it only loops through raw data set once. The old script looped through raw data for every stock.

### Challenges and Difficulties Encountered
There is always a risk of messing up code or more importantly the calculations when refactoring. My biggest fear is code running with no errors but a miss calculation in one of the outputs. I double checked outputs from both scripts to ensure all was working. One of the biggest limitations of refactored VBA is we have to hard code the tickers into an array. I would suggest making this dynamic by looping through the raw data first to collect all unique tickers and then looping through again to perform calculations.

## Results
I would recommend Steve use the refactored VBA script. Once he expands his universe of stocks and incorporates more years this script will vastly outperform original script. If Steve wishes to recommend any stocks from the list of 12 I would recommend two. ENPH and SEDG both had the highest average of returns across 2017/2018. However, I would recommend increasing his sample size to include more stocks and years.
