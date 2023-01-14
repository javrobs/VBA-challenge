# Multiple year stock data

##Introduction
This repository contains a VBA module that can analyse and summarize the opening and closing values of different stocks as well as the total volume of transactions made for each one over the years.

##Input file
In order for the code to work, the file may have the following format:
![Example file](/Screenshots/example.png)

##Instructions
1. This code should be imported into VBA through the VBA editor.  
2. Run under the developer tab clicking on the Macros button, or through the VBA editor run button. 

##Results
2018:
![This is the resulting image of running the code in the 2018](/Screenshots/2018.png)
2019:
![This is the resulting image of running the code in the 2018](/Screenshots/2018.png)
2020:
![This is the resulting image of running the code in the 2018](/Screenshots/2018.png)

##Justification
-Activating the worksheets was used over specifying the worksheet in every "cells" and "range" object for brevity and to minimize reworking the code from single sheet to a complete workbook.
-Initializing opening and volume variables to match second row of every sheet was the solution found 