Attribute VB_Name = "Module3"
Public Sub GetDatabaseLogFileLocations()
    Dim conn As Object, rs As Object
    Dim serverName As String, userName As String, password As String
    Dim sql As String
    Dim targetSheet As Worksheet
    Dim currentRow As Long
    
    ' Server and login details
    serverName = ""  ' Your server IP
    userName = ""  ' SQL server username
    password = ""  ' SQL server password

    ' Initialize the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=master;User ID=" & userName & ";Password=" & password & ";"

    ' SQL to fetch log file locations for each database
    sql = "SELECT DB_NAME(database_id) AS DatabaseName, physical_name AS LogFileLocation " & _
          "FROM sys.master_files " & _
          "WHERE type_desc = 'LOG'"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Prepare the target Excel sheet
    Set targetSheet = ThisWorkbook.Sheets("Log_File")  ' Change to your actual sheet name
    targetSheet.Range("A2:B" & targetSheet.Rows.Count).ClearContents

    ' Write headers
    targetSheet.Cells(1, 1).Value = "Database"
    targetSheet.Cells(1, 2).Value = "Log File Location"

    ' Start writing from row 2
    currentRow = 2

    ' Write the results to Excel
    While Not rs.EOF
        targetSheet.Cells(currentRow, 1).Value = rs.Fields("DatabaseName").Value  ' Database name
        targetSheet.Cells(currentRow, 2).Value = rs.Fields("LogFileLocation").Value  ' Log file location
        currentRow = currentRow + 1
        rs.MoveNext
    Wend

    ' Cleanup
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


