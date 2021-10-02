# DisplayData
Lightweight SQL query client
# Project Backgroun
Many times we just want to quickly query the database records, or delete, modify a record, use the SQL Server, ORACLE database platform of the client to open more slowly and resources, this project aims to meet the vast majority of developers a simple database query operations, the client tools is a lightweight database query,  The startup speed is fast and the resources are small.  
# The installation
Run setup.exe and follow the installation wizard to complete the installation.
Install the database access driver:
(1) DB2 support, requires a separate DB2 Run-time Client Lite installation
(2) Support Oracle, require separate installation of ODAC component (OLEDB)
(3) Support Sybase, Sybase OLEDB component is required to be installed (this software comes with the installation of this component)
(4) Support SQL Server, ADO component installation requirements (this software has the installation of the component)
# USE
# 1、General instructions for use  
(1) Single SQL statement execution  
You can submit a single SQL statement to the server for execution.  
(2) Execute multiple SQL statements at one time  
Multiple SQL statements can be isolated with Spaces or line breaks, and the program is submitted to the database server for execution at once.  (The prerequisite is grammar pass)  
(3) Multiple SQL statements are executed in sequence  
You can use semicolons (;) for multiple SQL statements.  ', the program will be submitted sentence by sentence, temporarily supports a maximum of 10 SQL statements.  
(4) Execute the specified SQL statement  
You can select some statements in the SQL edit box and run them. You can click the SQL command box twice to select the current line.  
(5) Execute SQL statement in transaction  
SQL statements that can be submitted are executed in a transaction, and unsuccessful programs are rolled back.  
(6) Export query results as Excel files  
You can export the query result set to a standard Excel file.  
(7) Button introduction  
Green triangle icon button: Executes the specified SQL statement.  
Green triangle icon button with Trans text: Executes the specified SQL statement in a transaction.  
Red Cross icon button: it only clears the content displayed in the interface, but does not delete the data in the database.  
Excel icon: Export the query result set to an Excel file.  
# 2、Data integrity description  
ServerType: Indicates the database service type, including MS SQL Server, Sybase ASE, and Oracle  
Server: indicates the database service name  
Server Port: the listening Port number of the database service. If you do not enter a Port number, the program uses the default Port number  
User: the User name used to access the data object  
Password: Password used to access the data object  
Database: The Database object to access, called Schema in ORACLE.  
# 3、Shortcut key description  
F1: Displays help information  
F2: Undo the edit operation in the SQL box  
F3: Restores edit operations in SQL boxes  
F5: Execute SQL statements  
F9: Displays all table objects in the selected database object  
F10: Adjust the size of the SQL command box to the maximum or restore it to the normal size  
# Principal Project manager  
xia miao  (totop@163.com)
