# VBA-challenge
Our code is fairly straight forward, I first start by making sure the code runs through all worksheets.
I then declare our variables with their best respective data type
Next I create our headers as shown in the example given, for columns I1:Q1 and O2:O4
I set the start of our TickerCount variable and j variable from row 2
I select a range for i, we briefly went over a LastRow variable example in class, however, when trying to add it to my script I couldnt get it to run properly.
So I decided to go with a simple range, in this case from row 2 to 800,000 which covers every single data point in all sheets
Next I used a formula that would loop through rows UNTIL column A(Ticker) does not equal the following rows ticker
when this happens pull that Ticker symbol and print it on column 9(I)
The following formula I got it from an example we did in class on 9/21, as directed by the professor on Sunday 10/2.
I'm not 100% confident on how to explain it, but it basically grabs the last rows value for column 6(F) for that respective ticker symbol and subtracts the first row value in column 3(C) for that same ticker symbol
this Yearly change is stored in column 10(J)
I then add a conditional format if a value is greater than or less than 0, with the help of the following link to select the color
https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
Next, I grab the percent change by using a simple math formula ((a-b)/b) and select the format type of the result as a "Percent"
Next I add the range of column 7(G) for all values for the respective ticker symbol.
Finally, I start working on the second half of my script, attempting the find the greatest increase, greatest decrease, and greatest volume Ticker symbol + value for that year(sheet)
Again, I select a range which fits all data points for column 11(K) and 12(L), in this case rows 2-10000
We loop through column 11(K) to find the biggest number, setting that number as our variable 'Greatest Increase" grabbing its respective Ticker symbol in column 9(I) and pasting it in column 16(P)
We again loop through column 11(K) to find the smallest number, setting that number as our variable 'Greatest Decrease" grabbing its respective Ticker symbol in column 9(I) and pasting it in column 16(P)
Lastly, we do the same as above for Greatest Volume
Finally we select to place our previous variables "Greatest Increase", "Greatest Decrease", and "Greatest Volume" and set them on their respective rows in column 17(Q)
Adjusting the format for "Greatest Increase" and "Greatest Decrease" as a Percentage.
Jujmp to the next worksheet and repeat.
