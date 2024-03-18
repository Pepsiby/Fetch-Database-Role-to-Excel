Attribute VB_Name = "Module2"
Public Sub RefreshUserDatabaseAccessSorted()
    Dim conn As Object, rsDB As Object, rsUser As Object
    Dim serverName As String, userName As String, password As String
    Dim dbSQL As String, userSQL As String, dbName As String
    Dim dbList() As String, dbCount As Long
    Dim i As Long, currentRow As Long
    Dim targetSheet As Worksheet
    Dim userData As Object, userList As Variant, key As Variant
    
    ' Initialize the dictionary to hold user-database mappings
    Set userData = CreateObject("Scripting.Dictionary")

    ' Server and login details
    serverName = ""  ' Your server IP
    userName = ""  ' SQL server username
    password = ""  ' SQL server password

    ' Initialize the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=master;User ID=" & userName & ";Password=" & password & ";"

    ' Get all database names
    dbSQL = "SELECT name FROM sys.databases WHERE state = 0 AND name NOT IN ('master', 'tempdb', 'model', 'msdb')"
    Set rsDB = conn.Execute(dbSQL)
    
    ' Count databases and populate the array
    While Not rsDB.EOF
        dbCount = dbCount + 1
        ReDim Preserve dbList(1 To dbCount)
        dbList(dbCount) = rsDB.Fields("name").Value
        rsDB.MoveNext
    Wend
    rsDB.Close

    ' Iterate through databases and collect user access
    For i = 1 To dbCount
        dbName = dbList(i)
        conn.DefaultDatabase = dbName

        userSQL = "SELECT dp.name AS DBUser FROM sys.database_principals dp WHERE dp.type IN ('S', 'U', 'G') AND dp.name NOT LIKE '##%'"
        Set rsUser = conn.Execute(userSQL)
        
        While Not rsUser.EOF
            Dim user As String
            user = rsUser.Fields("DBUser").Value
            
            If Not userData.Exists(user) Then
                Set userData(user) = CreateObject("Scripting.Dictionary")
            End If
            
            userData(user)(dbName) = True
            rsUser.MoveNext
        Wend
        rsUser.Close
    Next i

    conn.Close

        ' Prepare the target Excel sheet
    Set targetSheet = ThisWorkbook.Sheets("DB_Role")  ' Change to your actual sheet name
    
    ' Clear contents starting from row 2 to avoid deleting headers or other information in row 1
    targetSheet.Range("A2:B" & targetSheet.Rows.Count).ClearContents

    ' Start writing from row A2
    currentRow = 2

    ' Sort the usernames
    userList = userData.Keys
    Call QuickSort(userList, LBound(userList), UBound(userList))

    ' Write the sorted user-database pairs to Excel starting from row A2
    For Each key In userList
        Dim db As Variant
        For Each db In userData(key).Keys
            targetSheet.Cells(currentRow, 1).Value = key  ' Username
            targetSheet.Cells(currentRow, 2).Value = db   ' Database
            currentRow = currentRow + 1
        Next db
    Next key


    ' Cleanup
    Set conn = Nothing
    Set userData = Nothing
End Sub

' QuickSort function to sort the user list
Private Sub QuickSort(arr As Variant, first As Long, last As Long)
    Dim pivot As Variant, temp As Variant
    Dim i As Long, j As Long

    If first >= last Then Exit Sub

    pivot = arr((first + last) \ 2)
    i = first
    j = last

    While i <= j
        While arr(i) < pivot And i < last
            i = i + 1
        Wend
        While arr(j) > pivot And j > first
            j = j - 1
        Wend

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend

    If first < j Then QuickSort arr, first, j
    If i < last Then QuickSort arr, i, last
End Sub

