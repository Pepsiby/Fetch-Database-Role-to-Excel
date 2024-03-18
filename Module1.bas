Attribute VB_Name = "Module1"
Public Sub RefreshServerLoginsWithCreationDate()
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    Dim serverName As String
    Dim userName As String
    Dim password As String
    
    ' Update these details according to your SQL Server configuration
    serverName = ""  ' Your server IP
    userName = ""  ' Your SQL server username
    password = ""  ' Your SQL server password
    
    ' Connection String
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=master;User ID=" & userName & ";Password=" & password & ";"
    
    ' Open the connection
    conn.Open
    
    ' SQL to fetch server-level logins, their roles, and creation dates (year and month). Adjust according to your needs.
    Dim strSQL As String
    strSQL = "SELECT sp.name AS LoginName, sp.type_desc AS LoginType, " & _
             "CASE WHEN srm.role_principal_id IS NOT NULL THEN 'ALL' ELSE 'Limited' END AS ServerRole, " & _
             "FORMAT(sp.create_date, 'yyyy-MM') AS CreationDate " & _
             "FROM sys.server_principals sp " & _
             "LEFT JOIN sys.server_role_members srm ON sp.principal_id = srm.member_principal_id AND srm.role_principal_id = (SELECT principal_id FROM sys.server_principals WHERE name = 'sysadmin') " & _
             "WHERE sp.type IN ('S', 'U', 'G') AND sp.is_disabled = 0 AND sp.name NOT LIKE '##%'"
    
    ' Execute the SQL query
    rs.Open strSQL, conn
    
    ' Specify your Excel sheet and range where you want to put the data
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets("User")  ' Update with your actual sheet name
    
    ' Clear existing contents
    targetSheet.Range("A2:D1000").ClearContents  ' Adjusted to A2:D1000 to accommodate the new column
    
    ' Copy the data to Excel
    targetSheet.Range("A2").CopyFromRecordset rs
    
    ' Close the recordset and connection
    rs.Close
    conn.Close
    
    ' Clean up
    Set rs = Nothing
    Set conn = Nothing
End Sub

