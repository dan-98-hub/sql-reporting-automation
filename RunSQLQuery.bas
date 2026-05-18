' ── CONFIG ──────────────────────────────
Const SERVER   As String = "YOUR_SERVER"
Const DATABASE As String = "YOUR_DB"
' ────────────────────────────────────────

Sub RunQuery()
    On Error GoTo ErrorHandler

    Dim conn    As Object
    Dim rs      As Object
    Dim sql     As String
    Dim i       As Integer

    ' Get table name from user
    Dim tableName As String
    tableName = InputBox("Enter table name:", "SQL Query Runner")
    If tableName = "" Then Exit Sub

    ' Build connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & SERVER & _
                            ";Initial Catalog=" & DATABASE & _
                            ";Integrated Security=SSPI;"
    conn.Open

    ' Run query
    sql = "SELECT * FROM " & tableName
    Set rs = conn.Execute(sql)

    ' Clear old data
    Sheet1.Cells.Clear

    ' Write headers
    For i = 0 To rs.Fields.Count - 1
        Sheet1.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i

    ' Dump data
    Sheet1.Range("A2").CopyFromRecordset rs

    ' Format as table
    Dim tbl As ListObject
    Set tbl = Sheet1.ListObjects.Add( _
        xlSrcRange, Sheet1.UsedRange, , xlYes)
    tbl.Name = "SQLResults"
    tbl.TableStyle = "TableStyleMedium9"

    ' Timestamp
    Sheet1.Tab.Color = RGB(0, 176, 80)
    MsgBox "Done! Loaded data from: " & tableName & Chr(10) & _
           "Refreshed: " & Now()

    rs.Close
    conn.Close
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
End Sub
