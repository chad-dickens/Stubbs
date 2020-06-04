# Bank Statement Tracking

This VBA project guides users through the creation of a bank statement tracking file in Excel. Users are prompted to select two CSV files from their computer which are monthly bank statement exports. Provided these files are valid, the data will be inputed into the Excel file, the files will be exported to a macro free workbook and then emailed to another team. Various error handling scenarios are anticipated, with userforms and appropriate feedback to aid the user.

Main procedure is contained in MainModule. Common contains two sub procedures and a function called by MainModule. All other files are back end for the various userforms.

An example of how one of the three Excel tables looks is below:

![alt text](https://github.com/chad-dickens/Stubbs/blob/master/Bank%20Statement%20Tracking/TableExample.PNG)
