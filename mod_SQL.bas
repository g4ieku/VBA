Attribute VB_Name = "mod_SQL"
Option Explicit
' Module description: General SQL module

'https://learn.microsoft.com/en-us/sql/ado/reference/ado-api/datatypeenum?view=sql-server-ver16
'https://learn.microsoft.com/en-us/sql/ado/reference/ado-api/cursortypeenum?view=sql-server-ver16

Enum eState
    adStateOpen = 1         'for SQL late biding
    adOpenStatic = 3
End Enum
Enum adCommandTypeEnum      'for SQL late biding
    adCmdText = 1
End Enum
Enum adDateType             'for SQL late biding
    adInteger = 3
    adDouble = 5
    adDate = 7
End Enum



'Example of use
Sub ImportData()
    Dim file_downloadedDataFromSAP As String
    
    file_downloadedDataFromSAP = FilePath("Data From SAP")
    If file_downloadedDataFromSAP = "" Then Exit Sub
    
    Dim str_SQLquery As String                         'Store the query in a string
    
    str_SQLquery = "SELECT * FROM [Sheet1$]"
    'Call GetQueryResults(str_SQLquery, file_downloadedDataFromSAP, ws_downloadedData.Range("A2"), True)
    Call GetQueryResult(str_SQLquery, file_downloadedDataFromSAP, ThisWorkbook.Worksheets("FBL5N").Range("Q2"), True)
End Sub

''''SubName   : GetQueryResults
''''Purpose   : General sub, to connect to the closed workbook
Sub GetQueryResult(ByVal SQLquery As String, ByVal DataFilePath As String, destination As Range, Optional writeColumnsNames As Boolean = False, Optional headers As String = "YES")
    Dim rowCount As Long, i As Long
    
    'On Error GoTo eh
    Dim ADODBConnection As Object: Set ADODBConnection = CreateObject("ADODB.Connection")
    ADODBConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                         "Data Source=" & DataFilePath & ";" & _
                         "Extended Properties=""Excel 12.0;HDR=" & headers & ";"";"
              
    Dim ADODBRecordset As Object: Set ADODBRecordset = CreateObject("ADODB.Recordset")  'Run the query and store in a recordset
    ADODBRecordset.CursorType = adOpenStatic
    ADODBRecordset.Open SQLquery, ADODBConnection
    
    rowCount = ADODBRecordset.RecordCount
    destination.CopyFromRecordset ADODBRecordset                                        'Write data
    
    If writeColumnsNames = True Then                                                    'Write column names into row 1 of the worksheet
        For i = 0 To ADODBRecordset.Fields.Count - 1
            With ADODBRecordset.Fields(i)
                destination.Offset(-1, i).Value = .Name
            End With
        Next i
    End If
    
    ADODBRecordset.Close
    ADODBConnection.Close
    Set ADODBRecordset = Nothing
    Set ADODBConnection = Nothing
    
cleanup:                                                                                'Added code to close the recordset and connections correctly
    If Not (ADODBRecordset Is Nothing) Then
        If (ADODBRecordset.State And adStateOpen) = adStateOpen Then ADODBRecordset.Close
        Set ADODBRecordset = Nothing
    End If
    If Not (ADODBConnection Is Nothing) Then
        If (ADODBConnection.State And adStateOpen) = adStateOpen Then ADODBConnection.Close
        Set ADODBConnection = Nothing
    End If
    Exit Sub
eh:
    MsgBox Err.Description
    GoTo cleanup
End Sub


'Sub GetQueryResults(SQLQuery As String, DataFilePath As String, destination As Range, Optional writeColumnsNames As Boolean = False, Optional AddDatType As Boolean = False)
'    ' To add ADO reference select Tools->Reference and check "Microsoft ActiveX Data Objects Objects 6.1 Library"
'    Dim RowCount As Long, i As Long
'    Dim ws_name As String: ws_name = destination.Worksheet.Name
'
'    'On Error GoTo eh
'    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'              "Data Source=" & DataFilePath & ";" & _
'              "Extended Properties=""Excel 12.0;HDR=Yes;"";"
'
'    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")      ' Run the query and store in a recordset
'    rs.CursorType = adOpenStatic
'    rs.Open SQLQuery, conn
'
'    RowCount = rs.RecordCount
'
'    destination.CopyFromRecordset rs                                ' Write data
'
'    If writeColumnsNames = True Then                                 'Write column names into row 1 of the worksheet
'        For i = 0 To rs.Fields.Count - 1
'            With rs.Fields(i)
'                ThisWorkbook.Worksheets(ws_name).Range("A1").Offset(0, i).Value = .Name
'            End With
'        Next i
'
'        'Format the header row of the worksheet
'        With ThisWorkbook.Worksheets(ws_name).Range("A1").Resize(1, rs.Fields.Count)
'            .Interior.Color = rgbGreen
'            .Font.Color = rgbWhite
'            .Font.Bold = True
'            .EntireColumn.AutoFit
'        End With
'    End If
'
'
'    If AddDatType = True Then                                         'Loop through the recordset fields - Add Date Type
'        For i = 0 To rs.Fields.Count - 1
'            With rs.Fields(i)
'                'Debug.Print .Name, .Type
'
'                ThisWorkbook.Worksheets(ws_name).Range("A1").Offset(0, i).Value = .Name
'
'                If .Type = adDate Then
'                    ThisWorkbook.Worksheets(ws_name).Range("A1").Offset(1, i).Resize(RowCount).NumberFormat = "dd/mm/yyyy"
'                ElseIf .Type = adDouble And InStr(.Name, "$") Then
'                    ThisWorkbook.Worksheets(ws_name).Range("A1").Offset(1, i).Resize(RowCount).NumberFormat = "#,##0"
'                End If
'            End With
'        Next i
'    End If
'
''Added code to close the recordset and connections correctly
'cleanup:
'    If Not (rs Is Nothing) Then
'        If (rs.State And adStateOpen) = adStateOpen Then rs.Close
'        Set rs = Nothing
'    End If
'    If Not (conn Is Nothing) Then
'        If (conn.State And adStateOpen) = adStateOpen Then conn.Close
'        Set conn = Nothing
'    End If
'    Exit Sub
'eh:
'    MsgBox Err.Description
'    GoTo cleanup
'End Sub





Sub GetDataFromMultipleFiles()
    Dim conn As Object
    Dim rs As Object
    Dim SQLString As String
    Dim FilesPath As String
    Dim FileName As String
    Dim IsFirstFile As Boolean
    
    FilesPath = ThisWorkbook.Path & "\My Files\"
    FileName = Dir(FilesPath & "*.xlsx")
    
    Sheet1.Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    Set conn = CreateObject("ADODB.Connection")
    
    conn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & FilesPath & FileName & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    
    IsFirstFile = True
    Do Until FileName = ""
        'Debug.Print FileName
        
        If IsFirstFile Then
            SQLString = "SELECT * FROM [Sheet1$]"
            IsFirstFile = False
        Else
            SQLString = SQLString & " UNION ALL SELECT * FROM [Sheet1$] IN '" & _
                FilesPath & FileName & "' 'Excel 12.0 Xml;'"
        End If
        
        SQLString = SQLString & " WHERE [Oscar Nominations] > 0 AND [Genre] = 'Fantasy'"

        FileName = Dir
    Loop
    SQLString = SQLString & " ORDER BY [Title] ASC"
    
    'Debug.Print SQLString
    
    conn.Open

    Set rs = CreateObject("ADODB.Recordset")

    rs.ActiveConnection = conn
    rs.Source = SQLString
    rs.Open

    Sheet1.Range("A2").CopyFromRecordset rs

    rs.Close
    conn.Close
End Sub


''''SubName   : CreateSQLCommand
''''Purpose   : To export data to new file
Sub CreateSQLCommand()
    Call FolderCreate(folderGRIR)               'create folders to save the data there - on disk 'C:'
    
    Dim FullFilePath As String: FullFilePath = folderGRIR & "GRIR Comments " & Format(Now, "dd MMMM yyyy")
    On Error Resume Next
        Kill FullFilePath & ".xlsb"              'remove the file if already exists
    On Error GoTo 0
    
    Dim SQLCommand As String
    
    'Selection into a new workbook - workbook doesn't need to exist
    SQLCommand = _
        "SELECT [Action taken], [Communication level], [Action owner], Comments, [Responsible admin], Status " & _
            ",[Comment in local language], Deadline, [Last action date], [Feedback from OpCo], Key " & _
        "INTO [Output] " & _
            "IN '" & FullFilePath & ".xlsx' 'Excel 12.0 Xml;' " & _
        "FROM [Report$] "
    
    Call ExecuteSQLCommand(SQLCommand, FullFilePath)
End Sub

''''SubName   : ExecuteSQLCommand
''''Purpose   : To export data to new file
Sub ExecuteSQLCommand(ByVal SQLCommand As String, ByVal str_filePath As String)
    Dim conn As Object, cmd As Object
    
    'Create and open a connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                            "Data Source=" & ThisWorkbook.FullName & ";" & _
                            "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    
    On Error GoTo EndPoint                  'Try to open the connection, exit the subroutine if this fails
    conn.Open
    
    On Error GoTo CloseConnection           'If anything fails after this point, close the connection before exiting
    
    'Create the command object
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    cmd.CommandText = SQLCommand             'Use the command string that we passed into the procedure
    cmd.Execute                              'Try to execute the command
    conn.Close                               'Close the connection. This will happen anyway when the local variables go out of scope at the end of the subroutine
    
    'Free resources used by the command and connection. This will happen anyway when the local variables go out of scope at the end of the subroutine
    Set cmd = Nothing
    Set conn = Nothing
    
    Exit Sub                                'Exit here to make sure that the error handling code does not run
    
'ERROR HANDLERS. To close the connection safety if something happen

CloseConnection:                            'If the connection is opened successfully but a runtime error occurs later we end up here
    conn.Close
    Set conn = Nothing
    Debug.Print SQLCommand
    MsgBox Prompt:="An error occurred after the connection was established." & vbNewLine & vbNewLine & "Error number: " & Err.Number _
                & vbNewLine & "Error description: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Error After Connection Open"
           
    Exit Sub

EndPoint:                                   'If the connection failed to open we end up here
    MsgBox Prompt:="The connection failed to open." & vbNewLine & vbNewLine & "Error number: " & Err.Number & vbNewLine & "Error description: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Connection Error"
End Sub


''''SubName   : DivideList
''''Purpose   : Split table into separate worksheets/workbooks
Sub DivideList()
    Dim FilePath As String
    Dim ADODBConnection As Object
    Dim ADODBRecordset As Object
    Dim ADODBCommand As Object
        
    'ThisWorkbook.path & "/Data.xlsx"           'external workbook
    FilePath = ThisWorkbook.FullName
    
    Set ADODBConnection = CreateObject("ADODB.Connection")
    
    ADODBConnection.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & FilePath & ";" & _
        "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    
    ADODBConnection.Open
    
    Set ADODBRecordset = CreateObject("ADODB.Recordset")
    
    ADODBRecordset.ActiveConnection = ADODBConnection
    ADODBRecordset.Source = "SELECT DISTINCT [Customer No] FROM [NiceOutput$]"
    ADODBRecordset.Open
    
    Set ADODBCommand = CreateObject("ADODB.Command")
    ADODBCommand.ActiveConnection = ADODBConnection
    ADODBCommand.CommandType = adCmdText
    
    Do Until ADODBRecordset.EOF
        Debug.Print Len(ADODBRecordset.Fields("Customer No").Value)
        
        'create new worksheets in the same workbook
        'ADODBCommand.CommandText = _
            "SELECT * " & _
            "INTO [" & ADODBRecordset.Fields("Genre").Value & "] " & _
            "FROM [Film$] " & _
            "WHERE [Genre] = '" & ADODBRecordset.Fields("Genre").Value & "'"
        
        'create new worksheets in the closed workbook
        'ADODBCommand.CommandText = _
            "SELECT * " & _
            "INTO [" & ADODBRecordset.Fields("Genre").Value & "] " & _
                "IN '" & ThisWorkbook.Path & "\Output\All Genres.xlsx' 'Excel 12.0 Xml;' " & _
            "FROM [Film$] " & _
            "WHERE [Genre] = '" & ADODBRecordset.Fields("Genre").Value & "'"
        
        'create new worbooks
        ADODBCommand.CommandText = _
            "SELECT * " & _
            "INTO [" & ADODBRecordset.Fields("Customer No").Value & "] " & _
                "IN '" & ThisWorkbook.Path & "\Output\" & ADODBRecordset.Fields("Customer No").Value & ".xlsx' 'Excel 12.0 Xml;' " & _
            "FROM [NiceOutput$] " & _
            "WHERE CSTR([Customer No]) = Cstr('" & ADODBRecordset.Fields("Customer No").Value & "')"

        ADODBCommand.Execute
        
        ADODBRecordset.MoveNext
    Loop

    ADODBRecordset.Close
    ADODBConnection.Close
End Sub
