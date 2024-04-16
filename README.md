## How to auto populate a varying / arbitrary / unknown number of cells in Excel. Examples with and without VBA

<br/>

### Table of Contents
* [Summary](#Summary)
* [Detailed description](#Detailed-description)
* [Tested on](#Tested-on)
* [Walkthrough](#Walkthrough)
* [FAQ](#FAQ)


<br/><br/>
### Summary
Automatically populate cells based on a selection from a dropdown list.

The output can be of **various lengths** and repeated indefinitely:

![](https://github.com/JoeUgly/Excel_Auto_Populate/blob/master/assets/usage.gif)




<br/><br/>
### Detailed description
This repo has 2 different styles of worksheets available: VBA and non-VBA. Both styles achieve similar functionality, but use different methods.

1. VBA workbook: [spreadsheets/vba_example.xlsm](https://github.com/JoeUgly/Excel_Auto_Populate/raw/master/spreadsheets/vba_example.xlsm)\
     The autopopulation is activated by a keyboard shortcut.\
     Requires macros to be enabled.\
     Check the README sheet contained in the workbook.

2. Non-VBA workbooks:\
     MS Excel: [spreadsheets/non_vba_example.xlsx](https://github.com/JoeUgly/Excel_Auto_Populate/raw/master/spreadsheets/non_vba_example.xlsx)\
     LibreOffice Calc: [spreadsheets/non_vba_example.ods](https://github.com/JoeUgly/Excel_Auto_Populate/raw/master/spreadsheets/non_vba_example.ods)

     The autopopulation is activated automatically when an item from the dropdown list is selected.



<br/><br/><br/>
### Tested on
Microsoft Excel 2016 and LibreOffice Calc 7.1.4.2.

Since Excel uses an exclamation mark to designate a sheet name and LibreOffice Calc uses a period, I supplied two versions in the "spreadsheets" directory.



<br/><br/><br/>
### Walkthrough
All the following sections provide a walkthrough for recreating this workbook's functionality without VBA. However, **I strongly suggest using the workbooks I have provided** and modifying them to suit your needs, rather than starting from scratch.

This example uses 4 industries. Each industry has a unique list of parameters associated with it. The number of parameters in each list varies (between 3 and 14).

What we want is to select an industry from a dropdown list and have Excel automatically populate adjacent cells with the appropriate parameters for that industry. 

Our workbook will have 3 sheets. The first (named sample_log) will be where we select the industry from a dropdown list. The second sheet (named iu_param_table) will be where we store the table that contains all the industries and their parameters. The third sheet (named get_param) will be where we put all the logic of our formulas. The sample_log sheet will have cells referencing the get_param sheet, so that whatever result we get on the get_param sheet will also display on the sample_log sheet.



<br/><br/>
#### Create the industry and parameter table
To get started we are going to put all of the industries and parameters on a separate sheet, so that it is out of the way. We are going to name this sheet "iu_param_table".


<br/><br/>
#### Create the industry dropdown list
Now we are going to create a dropdown list using data validation. The list will consist of all of the industry names.

To create a dropdown list click on the Data section (on the ribbon), then Data Validation. Select list and the source. The source is where all the names of the industries are located in the workbook. Ex: iu_param_table!B2:E2



<br/><br/>
#### Create "get_param" sheet. The basic order of events for our formulas will be:
1. Get the industry name
2. Find that industry name in the table and return the column letter
3. Count the total number of rows (parameters) for that column
4. Count down (decrement) from step 3
5. Lookup the parameter using the number from step 4 and the column from step 2. 
<br/>
Many of these steps can be combined, but I chose to keep them separate in order to make it simpler to follow along.


<br/><br/>
#### 1. Get the industry name
Column name: Industry


This is an easy one. Get the industry from the sample_log sheet. If the corresponding cell from the sample_log sheet is empty, then leave blank.

`=IF(sample_log!B3="", "", sample_log!B3)`



<br/><br/>
#### 2. Find that industry name in the iu_param_table and return the column letter
Column name: Column Letter

Use the MATCH function and supply it with the industry name and the table from the sheet named "iu_param_table". The trailing zero specifies an exact match.
However, the MATCH function returns a number to designate the position in the table. Instead, we want the column letter for the sheet. By adding 65 to the result and converting it with the CHAR function, we get the appropriate column letter. If you don't start your table on column B then you will need to use a number different than 65.

If the cell from step 1 is blank, then repeat what is in the cell above.

`=IF(B3="", C2, CHAR(65+MATCH(B3, iu_param_table!$B$2:$E$2, 0)))`



<br/><br/>
#### 3. Count the total number of rows (parameters) for that column
Column name: Num of Params

Now we can count the number of items in that column. This will allow us to make our own FOR loop using something similar to a coordinate system. It will work like this:\
Column 3, row 3\
Column 3, row 4\
Column 3, row 5\
(until we reach a blank cell)

To do that, we supply the COUNTA function with a range of cells; the appropriate column starting at row 3 (that's where the data begins) until row 16 (where the data ends for the longest column in the table. If your data has more items, then increase this number).
This is what we are trying to achieve after the data has been evaluated:
`COUNTA(C3:C16)`

We are referencing column C in the above example from cell C3 (step 2) by using the INDIRECT function. Since the table is on another sheet we must also supply that by stating "iu_param_table!"

This is the formula we will actually use:

`=IF(sample_log!B3="", "", COUNTA(INDIRECT("iu_param_table!"&C3 & 3):INDIRECT("iu_param_table!"&C3 & 16)))`

<br/>
Now we have the number of parameters for that column.




<br/><br/>
#### 4. Count down from step 3
Column name: Remaining Num of Params

This step will be used to loop through all the parameters. We’ll start with the number supplied from the previous column and decrement it until we reach zero (no parameters remaining).

`=IF(E7="", "", IF(ISNUMBER(D8), D8, IF(E7-1>0, E7-1, "")))`




<br/><br/>
#### 5. Lookup the parameter using the number from step 4 and the column letter from step 2
Column name: Parameter

We are using the column letter supplied by cell C3 and the row number supplied by E3. We have to add 2 to it because the data in the table starts on row 3. Again, we are using the INDIRECT function to refer to the sheet named "iu_param_table".

`=IF(ISNUMBER(E3), INDIRECT("iu_param_table!"&C3&E3+2), "")`

And finally this should give a parameter for that industry. Use AutoFill to extend these formulas to the cells below. Each row will contain a different parameter for that industry, until no more remain. 




<br/><br/>
#### Display the results on the sample_log sheet
Paste this into the cell next to the industry that we selected on the "sample_log" sheet. Extend the formula into the rows below using the AutoFill feature. If your data on the sample_log sheet doesn't start on row 3 then adjust accordingly. 

`=get_param!F3`



<br/><br/>
Optional:\
You may notice that column C of sheet "get_param" has data extending down for all the rows that contain formulas. Unfortunately, there is no easy way of preventing this (due to circular logic errors) without making this example even more complicated. 

I supplied a full version of the workbook I personally use at "spreadsheets/others/complicated_example.xlsm". It contains different formulas, which prevent that column letter from repeating forever, it allows skipping rows, and many more features.




<br/><br/><br/>
### FAQ

#### How to add more columns to iu_param_table?

You must edit the formula in "Column Letter" (get_param sheet) to include more columns in the iu_param_table.\
Example (notice the "Z" near the end):\
`=IF(B3="", C1, CHAR(65+MATCH(B3, iu_param_table!$B$2:$Z$2, 0)))`

Note:\
The `CHAR + 65` method can only be used on single character column letters. Columns AA and beyond would need a new plan.


<br/><br/>
#### How to add more rows to iu_param_table?

You must edit the formula in "Num of Params" (get_param sheet) to include more rows in the iu_param_table.\
Example (notice the "50" near the end):\
`=IF(sample_log!B3="", "", COUNTA(INDIRECT("iu_param_table!"&C3 & 3):INDIRECT("iu_param_table!"&C3 & 50)))`



<br/><br/>
#### How to change the sheet names?

You must edit almost all of the formulas on the "get_param" sheet. Replace the default sheet names ("sample_log", "iu_param"table", etc) with new names.


<br/><br/>
#### How to change the column header names?

No additional steps are necessary.








<br/><br/>
