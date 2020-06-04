# Invoice Creation

The premise for this project is that you have a database of sales data and you need to create invoices for all of your customers based on this data. In this project, I have put my data into three tables in Microsoft Access and I then query these tables using SQL from VBA in Excel. When running the macro in Excel, the user is prompted to select their database file and if it is valid, they will be met with the following user form.

![alt text](https://github.com/chad-dickens/Stubbs/blob/master/Invoice%20Creation/Main_Window.PNG)

The run button is only enabled once the user has made a valid selection with no rates or address issues, preventing errors occuring while the process is running. The invoices are created from templates in two separate Excel worksheets, exported to PDFs, and then emailed to customers.
