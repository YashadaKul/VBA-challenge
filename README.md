# VBA-challenge
Uploading assignments for Module 2 of the Monash Bootcamp

Run the .vbs files in the following order:
1_ticker.vbs
2_StockVolume.vbs
3_YearlyChangePercentChange.vbs

Although my code for the yearly change and percent change worked for the sample file, it did not work on the main excel. My computer would crash. There is some redundancy in my code that I could not spot and it made the code too clunky for the large file.

Screenshot provided is for the sample file, alphabetical_testing_1.xlsm
 
The codes here work for the sample file but some assumptions had to be made:
1. If the >ticker< column did not have the names in alphabetical order, then they would have to be sorted alphabetically first before running the code. The code logic is meant to pick cells if they dissimilar, so the ticker names would have to be sorted first before running the code. In this case, they were already sorted, but for some other file, the code wont be right if the ticker names are scatterred in the column. 
2. The values for stock volume should align with the ticker names, in this case they did but its not necessary that they would for some other file. So, ideally, the stock volume and ticker name should be 1 sub where both are calculated and printed together. 



