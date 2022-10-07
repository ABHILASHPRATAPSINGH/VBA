Attribute VB_Name = "Module1"
Option Explicit

Sub CreateSQLQuery()

    Dim sqlquery As String
    
    'SQLQuery = _
        "SELECT " & _
            " [f].[Title] AS [Film Name]" & _
            ",[f].[Run Time] AS [Length]" & _
            ",[f].[Release Date]" & _
            ",[f].[Oscar Wins] " & _
        "FROM " & _
            "[Film$] AS [f]"
    sqlquery = "select * from [FilmYears$A15:D26]"
    
    'Run the query with the SQL string
    GetQueryResults sqlquery
    
End Sub

Sub GetQueryResults(sqlquery As String)

    Dim movieFilePath As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim i As Integer
    Dim RowCount As Long, ColCount As Long
    
    'Exit the procedure if no query was passed in
    If sqlquery = "" Then
        MsgBox _
            Prompt:="You didn't enter a query", _
            Buttons:=vbCritical, _
            Title:="Query string missing"
        Exit Sub
    End If
    
    'Check that the Movies workbook exists in the same folder as this workbook
    movieFilePath = ThisWorkbook.Path & "\Movies.xlsx"
    
    If Dir(movieFilePath) = "" Then
        MsgBox _
            Prompt:="Could not find Movies.xlsx", _
            Buttons:=vbCritical, _
            Title:="File not found"
        Exit Sub
    End If
    
    'Create and open a connection to the Movies workbook
    Set cn = New ADODB.Connection
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & movieFilePath & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=NO';"
    
    'Try to open the connection, exit the subroutine if this fails
    On Error GoTo EndPoint
    cn.Open
    
    'If anything fails after this point, close the connection before exiting
    On Error GoTo CloseConnection
    
    'Create and populate the recordset using the SQLQuery
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorType = adOpenStatic
    
    rs.Source = sqlquery    'Use the query string that we passed into the procedure
    
    'Try to open the recordset to return the results of the query
    rs.Open
    
    'If anything fails after this point, close the recordset and connection before exiting
    On Error GoTo CloseRecordset
    
    'Get count of rows returned by the query
    RowCount = rs.RecordCount
    
    'Exit the procedure if no rows returned
    If RowCount = 0 Then
        MsgBox _
            Prompt:="The query returned no results", _
            Buttons:=vbExclamation, _
            Title:="No Results"
        Exit Sub
    End If
    
    'Get the count of columns returned by the query
    ColCount = rs.Fields.Count
    
    'Create a new worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    'Select the worksheet to avoid the formatting bug with CopyFromRecordset
    ThisWorkbook.Activate
    ws.Select
    
    'Format the header row of the worksheet
    With ws.Range("A1").Resize(1, ColCount)
        .Interior.Color = rgbCornflowerBlue
        .Font.Color = rgbWhite
        .Font.Bold = True
    End With
    
    'Copy values from the recordset into the worksheet
    ws.Range("A2").CopyFromRecordset rs
    
    'Write column names into row 1 of the worksheet
    For i = 0 To ColCount - 1
        With rs.Fields(i)
            ws.Range("A1").Offset(0, i).value = .Name
            
            'Apply a custom date format to date columns
            If .Type = adDate Then
                ws.Range("A1").Offset(1, i).Resize(RowCount, 1).NumberFormat = "dd mmm yyyy"
            End If
        End With
    Next i
    
    'Change the column widths on the worksheet
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    'Close the recordset and connection
    'This will happen anyway when the local variables go out of scope at the end of the subroutine
    rs.Close
    cn.Close
    
    'Free resources used by the recordset and connection
    'This will happen anyway when the local variables go out of scope at the end of the subroutine
    Set rs = Nothing
    Set cn = Nothing
    
    'Exit here to make sure that the error handling code does not run
    Exit Sub
    
'========================================================================
'ERROR HANDLERS
'========================================================================
CloseRecordset:
'If the recordset is opened successfully but a runtime error occurs later we end up here
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
    Debug.Print sqlquery
    
    MsgBox _
        Prompt:="An error occurred after the recordset was opened." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Error After Recordset Open"
    
    Exit Sub

CloseConnection:
'If the connection is opened successfully but a runtime error occurs later we end up here
    cn.Close
    
    Set cn = Nothing
    
    Debug.Print sqlquery
    
    MsgBox _
        Prompt:="An error occurred after the connection was established." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Error After Connection Open"
    
    Exit Sub
    
'If the connection failed to open we end up here
EndPoint:
    MsgBox _
        Prompt:="The connection failed to open." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Connection Error"
    
End Sub

Sub DeleteAllButMenuSheet()

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is MenuSheet Then ws.Delete
    Next ws
    
    Application.DisplayAlerts = True
    
End Sub
Sub GetQueryResult(sqlquery As String, input_Filepath As String, output_Filepath As String, output_worksheet As String, Optional ByVal newSheet As Boolean = True, Optional ByVal output_Range As String = "A1")
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim i As Integer
    Dim RowCount As Long, ColCount As Long
    
    'Exit the procedure if no query was passed in
    If sqlquery = "" Then
        MsgBox _
            Prompt:="You didn't enter a query", _
            Buttons:=vbCritical, _
            Title:="Query string missing"
        Exit Sub
    End If
    
    'Check that the Movies workbook exists in the same folder as this workbook
    'input_Filepath = ThisWorkbook.Path & "\Movies.xlsx"
    
    If Dir(input_Filepath) = "" Then
        MsgBox _
            Prompt:="Could not find Movies.xlsx", _
            Buttons:=vbCritical, _
            Title:="File not found"
        Exit Sub
    End If
    
    'Create and open a connection to the Movies workbook
    Set cn = New ADODB.Connection
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & input_Filepath & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    
    'Try to open the connection, exit the subroutine if this fails
    On Error GoTo EndPoint
    cn.Open
    
    'If anything fails after this point, close the connection before exiting
    On Error GoTo CloseConnection
    
    'Create and populate the recordset using the SQLQuery
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorType = adOpenStatic
    
    rs.Source = sqlquery    'Use the query string that we passed into the procedure
    
    'Try to open the recordset to return the results of the query
    rs.Open
    
    'If anything fails after this point, close the recordset and connection before exiting
    On Error GoTo CloseRecordset
    
    'Get count of rows returned by the query
    RowCount = rs.RecordCount
    
    'Exit the procedure if no rows returned
    If RowCount = 0 Then
        MsgBox _
            Prompt:="The query returned no results", _
            Buttons:=vbExclamation, _
            Title:="No Results"
        Exit Sub
    End If
    
    'Get the count of columns returned by the query
    ColCount = rs.Fields.Count
    
    'Create a new worksheet
    
    Workbooks.Open (output_Filepath)
    
    If newSheet = True Then
        Set ws = Worksheets.Add
    Else
        Set ws = Worksheets(output_worksheet)
    End If
    
    'Select the worksheet to avoid the formatting bug with CopyFromRecordset
    'ThisWorkbook.Activate
    ws.Select
    
    'Format the header row of the worksheet
    With ws.Range(output_Range).Resize(1, ColCount)
        .Interior.Color = rgbCornflowerBlue
        .Font.Color = rgbWhite
        .Font.Bold = True
    End With
    
    'Copy values from the recordset into the worksheet
    ws.Cells(Range(output_Range).Row + 1, Range(output_Range).Column).CopyFromRecordset rs
    
    'Write column names into row 1 of the worksheet
    For i = 0 To ColCount - 1
        With rs.Fields(i)
            ws.Range(output_Range).Offset(0, i).value = .Name
            
            'Apply a custom date format to date columns
            If .Type = adDate Then
                ws.Range(output_Range).Offset(1, i).Resize(RowCount, 1).NumberFormat = "dd mmm yyyy"
            End If
        End With
    Next i
    
    'Change the column widths on the worksheet
    ws.Range(output_Range).CurrentRegion.EntireColumn.AutoFit
    
    'Close the recordset and connection
    'This will happen anyway when the local variables go out of scope at the end of the subroutine
    rs.Close
    cn.Close
    
    'Free resources used by the recordset and connection
    'This will happen anyway when the local variables go out of scope at the end of the subroutine
    Set rs = Nothing
    Set cn = Nothing
    
    'Exit here to make sure that the error handling code does not run
    Exit Sub
    
'========================================================================
'ERROR HANDLERS
'========================================================================
CloseRecordset:
'If the recordset is opened successfully but a runtime error occurs later we end up here
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
    Debug.Print sqlquery
    
    MsgBox _
        Prompt:="An error occurred after the recordset was opened." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Error After Recordset Open"
    
    Exit Sub

CloseConnection:
'If the connection is opened successfully but a runtime error occurs later we end up here
    cn.Close
    
    Set cn = Nothing
    
    Debug.Print sqlquery
    
    MsgBox _
        Prompt:="An error occurred after the connection was established." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Error After Connection Open"
    
    Exit Sub
    
'If the connection failed to open we end up here
EndPoint:
    MsgBox _
        Prompt:="The connection failed to open." & vbNewLine _
            & vbNewLine & "Error number: " & Err.Number _
            & vbNewLine & "Error description: " & Err.Description, _
        Buttons:=vbCritical, _
        Title:="Connection Error"
    

End Sub





