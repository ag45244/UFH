# Utility File Handling Python Job

This python script automates daily Business Analyst work by handling the transfer and manipulation of electronic data.
*can be modified within any other organization to fit their business needs

The following bullets are the main methods involved for this process:
•File Verification
•Modify-this manipulating edi, txt, csv or excel files. This can include combining 100's of files into one or renaming them all then archiving them.
•PGP Encrypt/Decrypt-Some data is meant to stay confidential which requires encrypting data when going out or decrypting it when inbound
•File Transfer Protocol (FTP) Put/Get
•Archive
•Email Notification w/(o) attachment

Python Modules Required:
pyodbc, os, re, shutil, glob, ftplib, datetime, smtplib, email, subprocess

Technology at play: 
The SQL code to create the UtilityFileHandlingReference table has been provided. This table stores information regarding  SFTP credentials, transfer status (inbound or outbound), PGP requirements, and a schedule for when a specific row should run.

By using the python ODBC (open database connectivity) module, I can use python to communicate with MS SQL Server and interact with the data from the dbo.UtilityFileHandlingReference table. With python, the logic is created and the SQL can be embedded.
