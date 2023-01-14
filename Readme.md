# Multiple year stock data

## Introduction
This repository contains a VBA module that can summarize stock values over different tabs by creating two tables:
1. The yearly change of each stock, the percentage of change and the total volume of transactions.
2. The ticker with the best increase percentage, the worst decrease percentage, and the most transactions.

## Input file
In order for the code to work, the file may have the following format:
![Example file](/Screenshots/example.png)
- Before running the macro, the stock transactions should be ordered alphabetically and the transactions must be ordered chronologically.
- The first table will populate the "I:L" columns. The second table will populate "O1:P4". Make sure these cells are empty to ensure data isn't lost.

## Instructions
1. This code should be imported into VBA through the VBA editor.  
2. Run under the developer tab clicking on the Macros button, or through the VBA editor run button. The sub is called summary.
3. Wait for the code to finish running.

## Results
2018:
![This is the resulting image of running the code in the 2018](/Screenshots/2018.png)
2019:
![This is the resulting image of running the code in the 2018](/Screenshots/2018.png)
2020:
![This is the resulting image of running the code in the 2018](/Screenshots/2018.png)

## Justification
- `.activate` was used over specifying the worksheet in every `cells` and `range` for brevity and to minimize reworking the code from single sheet to a complete workbook.
- The `for` loop for the first table starts at 3 because starting at 2 would cause the conditional to compare to the header row.
- The conditional formatting was only applied to the "Yearly Change" values to match the challenge examples.
