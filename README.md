# Production_Management

This program was written in Python 3.5.2 for windows environments. You MUST enter your own SQL database information in the fields with 7 
astersisks, or it won't do you much good to run any of these.

High level overview:

These programs are meant to work together to form the skeleton of a production management system. The Label_Program prints a label that is 
meant to be used as a production ticket. It should display all the information the production line needs to produce the part. 
Reporter_Program is meant to report results anywhere on the production line from a program that outputs .csv files. Embosser_Program takes 
the information from the barcode on the label that the Label_Program outputs and embosses that onto a plate.

Slightly lower level overview: 

-The Label_Program was made to print up production labels from a specific SQL database that you will have to adapt to suit your needs.
This program starts by querying the database to try to determine if this is a reprint. It then moves to insert the information. It prints 
by default to a Zebra 300 dpi label printer and is formatted specifically for that. Finally, it uses a redundant commit to the SQL 
database.

-The Reporter_Program looks for changes in a directory every second. It then parses to the end of the changed file (.csv) and reports it's 
content to a SQL database. The program has been customized to work with a specific machine that doesn't change it's UID value (Serial 
Number) until the test passes. This means that it doesn't push every new entry but instead it holds a value until a different value is 
presented. This is a very specific program that should be adapted before most uses.

-The Embosser_Program takes input from the label and sends it both to an embosser and then the specified SQL database. The process is 
twofold. First, the embossing needs to happen no matter how the SQL commit happens (or doesn't) and second, there are backup commit 
statements that mirror those in the Label_Program. The embossing program *should* work with most embossers, as the syntax is similar if
not identical for most products.

