# VBA-challenge
VBA Challenge Homework Assignment
Anthony Gonzalez
Southern Methodist University
Data Science Bootcamp


--  Purpose and Instruction
The purpose of this homework assignment is to demonstrate my understanding and of, and ability to code for, the four programming fundamentals in VBA: Conditionals, Iterations, Functions, Variables/Arrays.

For this homework assignment, the instructor provided an excel workbook with daily stock data for 2014, 2015, and 2016; with data for each year on a separate sheet. Data points include ticker symbol, date, open price, high price, low price, close price, and trading volume for each stock, each day. 

The assignment called for a summarized list of all stocks, the yearly change, the yearly change as a percentage, the total stock volume, and conditional formatting for yearly change. 

-- Summary of Approach
To create this summarized list, I created an Excel VBA macro that performed the following tasks.

1. Create a loop so that a stock summary table would be created for each workseet.
2. Use Range cell references to create the stock summary table header row
3. Find the last row in each list. Each year had a varying number of row data.
4. For each list: A loop and conditional is used to determine the first and last instance of each ticker symbol.
5. Using this loop and conditional, the macro is able to determine the first open price, the last close price, and the sum of daily stock volume for each ticker.
6. These metrics are transposed to the summary table using dynamic cell references and variables. 
7. A separate loop and conditional statement is used to fill each cell green or red; dependent on yearly change.


