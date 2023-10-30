# VBA-challenge
VBA Homework assignment

I uploaded two file types for the code depending on which is easier for grading purposes. Both files are otherwise the same "Yearly_stock_data_summary.cls" & "Yearly_stock_data_summaryVBA".

Included .png files in this VBA-challenge repository show results from each worksheet of the excel file, after the VBA code was run.

The VBA script runs through all worksheets at once. There are ways to format the code where it will only run through the active worksheet, but this can result in errors if you cycle through the worksheets while the code is running, so I decided to loop it through all worksheets in the workbook by calling each worksheet individually (as opposed to re-running the VBA code on each worksheet). 

For the conditional formatting, I did not include coding for values that remained the same from the beginning to the end of the year. The instructions did not specify this point, so I left them un-colored. (simply adding the greater than or equals to notation ">=" is the only change needed to apply that functionality). 

I worked on this project alone. I used StackOverflow to look at the various examples of how to loop through worksheets. 
