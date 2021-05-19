# Minibase
A personal project to help keep track of miniatures that I have purchased or will recieve. It was created using Google Sheets, Microsoft Azure SQL Database, Microsoft SQL Server Management Studio 18 (SSMS), and C# so I could get a little more comfortable with each one. 

# How it was made.
I took data from 2 different Google Sheets that had information about miniatures for specific Kickstarters (listed below) and combined them together. I deleted columns that I found irrelevant and the remaining columns became the basis of the Miniature Table in the Azure SQL Database. I used SSMS to create my Miniature table and created another table called Company. A foreign key links the two together. I then went to Google's Cloud Platform and obtained the necessary credenetials to use the Google Sheets API. Using C#, I connected to the Azure SQL Database and was able to take the information from the Google Sheet to fill my Miniature table up by writing the querys in the C# program.

#Google Sheets
Blacklist Miniatures Kickstarters 1 + 2 : https://docs.google.com/spreadsheets/d/1gge4QeqrcKawTBkuUW5XlH2qzDDdtQXZkXti55oSocs/edit#gid=0
Reaper Bones 5 Kickstarter : https://docs.google.com/spreadsheets/d/1Tza7h0aj-icHh-rD2z5RM6E-DQmMqR3g/edit?rtpof=true

My Google Sheet: https://docs.google.com/spreadsheets/d/192piBgNM8YBSnT5SOGHp8bmoE0kvDMFRvt-7HM7cQKQ/edit?usp=sharing
