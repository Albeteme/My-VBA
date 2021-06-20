# My-VBA
This VBA script is used to process real stock market data in a Microsoft Excel workbook.

## Background
In this project, I created a VBA (Visual Basic) script to analyze some stock market data. The data comes from a Microsoft Excel workbook with stock data of three years (2014, 2015, and 2016). Each year is a different tab/sheet inside the workbook. 

The Excel workbook being too large to upload, I have the vba script and the excel document separately loaded to this GitHub repository. 
You can find the original files used to run this script here:

https://gt.bootcampcontent.com/GT-Coding-Boot-Camp/gt-atl-data-pt-03-2021-u-c/tree/master/02-VBA-Scripting/Homework/Instructions

## Test running

I used excel to run this script on both the testing Excel workbook (alphabetical_testing.xlsx) and on the final multiple year stock workbook (multiple_year_stock_data.xlsx).

## The Script

The script may take few minutes to run because it goes through every single sheet. Just be patient.

![image](https://user-images.githubusercontent.com/75787486/122659819-ad956900-d149-11eb-822e-75dffd382af4.png)

As the script runs, it is doing the following:

* It loops through all the stocks for one year for each run and takes the following information:

  * The ticker symbol

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  
  * The total stock volume of the stock.

* It applies conditional formatting by highlighting positive yearly change values in green and negative yearly change values in red.

* Finally, it return the stock with the greatest percent increase, greatest percent decrease, and greatest total volume

![image](https://user-images.githubusercontent.com/75787486/122659821-b4bc7700-d149-11eb-96f0-d86bd2a278cb.png)

After the script has finished running, the results can be seen in the Excel workbook. 
The screenshoots show the yearly change for 2014,2015 and 2016. 
