# SECDataParseVBA
Download and query SEC XBRL tags from quarterly and annual financial data

<b>Purpose</b>
SEC fundamental data has historically been a manual process to view and obtain. You typically have two options, (1) find the filing directly through Edgar, or (2) use a web application which may or may not have a paywall.

In recent years XBRL has made strides in standardizing SEC filing data. Such to the extent that anyone can directly download all filing XBRL tag data directly from the SEC.gov website. 

My goal in this project is to make data more accessible in the tool that millions of financial analyst around the world use everyday (excel). A simple set up process will allow users with rudimentry VBA skills to access SEC data programatically. From there UDF's will allow even non-VBA excel users to quickly obtain any peice of SEC filing data without ever having to leave excel.

Note: This repository is in-progress. As so far the class being built allows you to download quarterly data from the SEC and unzip located at the following directory: C:\SECVba. A second class can be utilized to load the data into SQL server. Progress is being made on a third class with methods to return SQL strings with more simple user inputs (think =returnSECDataArray(Array(adsh, ddate, value), currentassets))) 

Without the third class queries can still be made to SQL server but are not as user friendly. Refer to the example module in this respository for an example query.


<b>Requirements:</b>
<br>
SQL Server (2017) and related driver: msoledbsql_18.1.0.0_x64.msi(or msoledbsql_18.1.0.0_x84.msi)

<b>Instructions:</b>
<br>
<b>(1)</b> Place the classes into a VBA project
<br>
<b>(2)</b> Download SQL Server and the msi file for the ability to work with SQL server from VBA
<br>
<b>(3)</b> Place "SECVba" file at the C:\ directory"
<br>
<b>(4)</b> Use the class methods in a module to download SEC data, create SQL Server tables, and load the tables for the data you've downloaded. 
<br>
Note: It's highly suggested you include data validation of some sort for your loads. For example, after loading Q3'2018 num tables you should note that somewhere either as a record in a SQL table, or a sheet in excel. Purpose of this is to ensure you don't load duplicates. 



Example use to instantly query array of data into excel from SQL server (current assets for selected SEC filers as of 20180331):
![alt text](https://github.com/dylan1218/SECDataParseVBA/blob/master/ExampleArrayResult.PNG)
