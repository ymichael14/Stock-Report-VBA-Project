# VBA Project: The VBA of Wall Street

## Background

A reporting analyst may have multiple pages that need the same work done. With Excel VBA, you can do that and automate the process. Using test data, I was able automate and do calculations and have those calculations appear in my worksheets in excel. Valuable tool to automate mundane tasks on similarly formatted excel files.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock Market Analyst

Created a script that loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

**Note:** Used conditional formatting that will highlight positive change in green and negative change in red.

The result should match the following image:

![moderate_solution](Images/moderate_solution.png)

## Bonus features added

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

![hard_solution](Images/hard_solution.png)

Made the appropriate adjustments to my VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.

##Considerations

* Make sure that the script acts dynamically, so code acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with one click of a button.



## Screenshots of excel files after running the code are in the repository


