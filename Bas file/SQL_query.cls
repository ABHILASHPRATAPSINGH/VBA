VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQL_query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Sub test1()
'    Dim getSQL As New SQL_query
'
'    '$ should be inserted b/w sheetname of cellrange
'    Query = "Select * from [Sheet1$a1:c11] where Fruits='Bannana'"
'
'    Dim wbp As Workbook
'    Set wbp = ThisWorkbook
'
'    wbc = ThisWorkbook.path & "\" & ThisWorkbook.Name
'
'    Call getSQL.GetQueryResults(Query, wbc, ThisWorkbook, "Sheet1", "O5", True)
'
'
'End Sub


Sub GetQueryResults(ByVal SQLQuery As String, _
                                    ByVal wbCopy_path As String, _
                                    wbPaste As Workbook, _
                                    ByVal wk_paste As String, _
                                    cellName, _
                                    Optional pasteHeader As Boolean = True)

    Dim MovieFilePath As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ws As Worksheet
    Dim i As Integer
    Dim RowCount As Long, ColCount As Long
    

     'Exit the procedure if no query was passed in
'    If SQLQuery = "" Then
'        MsgBox _
'            Prompt:="You didn't enter a query", _
'            Buttons:=vbCritical, _
'            Title:="Query string missing"
'        Exit Sub
'    End If
    
    'Check that the Movies workbook exists in the same folder as this workbook
    'MovieFilePath = "C:\Users\DELL\Documents\VBA\Movies.xlsx"
    
'    If Dir(MovieFilePath) = "" Then
'        MsgBox _
'            Prompt:="Could not find Movies.xlsx", _
'            Buttons:=vbCritical, _
'            Title:="File not found"
'        Exit Sub
'    End If
    
    'Create and open a connection to the Movies workbook
    Set cn = New ADODB.Connection
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & wbCopy_path & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=Yes';"
    
    'Try to open the connection, exit the subroutine if this fails
    On Error GoTo EndPoint
    cn.Open
    
    'If anything fails after this point, close the connection before exiting
    On Error GoTo CloseConnection
    
    'Create and populate the recordset using the SQLQuery
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorType = adOpenStatic
    
    rs.Source = SQLQuery    'Use the query string that we passed into the procedure
    
    'Try to open the recordset to return the results of the query
    rs.Open
    
    'If anything fails after this point, close the recordset and connection before exiting
    On Error GoTo CloseRecordset
    
    'Get count of rows returned by the query
    RowCount = rs.RecordCount
    
    'Exit the procedure if no rows returned
    If RowCount = 0 Then
'        MsgBox _
'            Prompt:="The query returned no results", _
'            Buttons:=vbExclamation, _
'            Title:="No Results"
'        Exit Sub
    End If
    
    'Get the count of columns returned by the query
    ColCount = rs.Fields.Count
    
    'Create a new worksheet
    'Set ws = ThisWorkbook.Worksheets.Add
    
    'Select the worksheet to avoid the formatting bug with CopyFromRecordset
    wbPaste.Activate
    Set ws = wbPaste.Worksheets(wk_paste)
    ws.Select
    
    rownum = Range(cellName).Row
    columnNum = Range(cellName).Column
    
    'Format the header row of the worksheet
'    With ws.Range("A1").Resize(1, ColCount)
'        .Interior.Color = rgbCornflowerBlue
'        .Font.Color = rgbWhite
'        .Font.Bold = True
'    End With
    
    'Copy values from the recordset into the worksheet
    If pasteHeader Then
        ws.Cells(rownum + 1, columnNum).CopyFromRecordset rs
    
    
    'Write column names into row 1 of the worksheet
    For i = 0 To ColCount - 1
        With rs.Fields(i)
            
            ws.Cells(rownum, columnNum).Offset(0, i).Value = .Name
            ws.Range("A1").Offset(0, i).Value = .Name

            'Apply a custom date format to date columns
'            If .Type = adDate Then
'                ws.Range("A1").Offset(1, i).Resize(RowCount, 1).NumberFormat = "dd mmm yyyy"
'            End If
        End With
    Next i
    Else
        ws.Cells(rownum, columnNum).CopyFromRecordset rs
    End If
    
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
    
    Debug.Print SQLQuery
    
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
    
    Debug.Print SQLQuery
    
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



