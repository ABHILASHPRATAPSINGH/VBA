VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Sub test1()
''? how to check all sheet name
''Library: active x object library 6.0
''Advantages: without option worksheet we can all sheet name.
''parameter:-filepath:- complete filepath, sheetArr:- 1d array which contains sheet name that need to be checked.
''output: it return nonExist sheet name from array
'
'    Dim pathLibCls As New sheetExist
'    Dim sheetArr() As Variant
'
'    sheetArr = Array("Sheet1", "Sheet5", "sdfsd")
'
'    MsgBox pathLibCls.nonExistSheetName("C:\Users\DELL\Desktop\variableDeclare.xlsm", sheetArr)
'End Sub
'
'Sub test2()
''? how to check all sheet name
''Library: active x object library 6.0
''Advantages: without option worksheet we can all sheet name.
''parameter:-filepath:- complete filepath, sheetArr:- 1d array which contains sheet name that need to be checked.
''output: it return nonExist sheet name from array
'
'    Dim pathLibCls As New sheetExist
'    Dim sheetArr() As Variant
'
'    sheetArr = Array("Sheet1", "Sheet5", "sdfsd")
'
'    MsgBox pathLibCls.nonExistSheetName("C:\Users\DELL\Desktop\variableDeclare.xlsm", sheetArr)
'
'
'End Sub














Public Function nonExistSheetName(filePath, array_sheetName) As String
'? how to check all sheet name
'Library: active x object library 6.0
'Advantages: without option worksheet we can all sheet name.
'Parameter:-filepath:- complete filepath, sheetArr:- 1d array which contains sheet name that need to be checked.
'output: it return nonExist sheet name from array
    
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    Dim col() As String
    
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & filePath & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=Yes';"
    
    cn.Open
    
    'Set rssheets = cn.OpenSchema(adschemaTables)
    Set rssheets = cn.OpenSchema(20)
    
    i = 0
    Do Until rssheets.EOF
        val1 = Replace(rssheets.Fields("Table_Name").Value, "$", "")
        val1 = Replace(val1, "'", "")
        ReDim Preserve col(i)
        col(i) = val1
        i = i + 1
        rssheets.movenext
    Loop
    
    For Each sht_item In array_sheetName
        If Not IsNumeric(Application.Match(sht_item, col, 0)) Then
            nonExistSheetName = sht_item
            Exit For
        End If
    Next sht_item
    
    rssheets.Close
    cn.Close
    Set cn = Nothing
    Set rssheets = Nothing
End Function

Public Function checkSheetExist(filePath As String, sheetName As String) As Boolean
'? how to check all sheet name
'Library: active x object library 6.0
'Advantages: without option worksheet we can all sheet name.
'Parameter:-filepath:- complete filepath, sheetName: sheet name that need to check
'output: it return boolean value, true if sheet exist otherwise false.
    
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    Dim col() As String
    
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & filePath & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=Yes';"
    
    cn.Open
    
    Set rssheets = cn.OpenSchema(20)
    
    'Set rssheets = cn.OpenSchema(adschemaTables)
    
    i = 0
    Do Until rssheets.EOF
        val1 = Replace(rssheets.Fields("Table_Name").Value, "$", "")
        val1 = Replace(val1, "'", "")
        ReDim Preserve col(i)
        col(i) = val1
        i = i + 1
        rssheets.movenext
    Loop
    
    checkSheetExist = False
        If IsNumeric(Application.Match(sheetName, col, 0)) Then
            checkSheetExist = True
        End If
        
    rssheets.Close
    cn.Close
    Set cn = Nothing
    Set rssheets = Nothing
End Function


