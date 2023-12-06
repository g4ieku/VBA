Attribute VB_Name = "mod_Utylities"
Option Explicit

'FunctionName  : DeleteOutputFiles
''''''Purpose  : Deleting all files in the folder
Function DeleteOutputFiles(ByVal Path As String)
    If Dir(Path & "*.xlsx") <> "" Then Kill Path & "*.xlsx"
End Function


Sub su3_transaction()
    'Go to the transaction su3 and change spool control
    With Session
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nsu3"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
        .findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/ctxtSUID_ST_NODE_DEFAULTS-SPLD").Text = "LOCL"      'OutputDevice
        .findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/chkSUID_ST_NODE_DEFAULTS-SPDA").Selected = False    'Delete After Output
        .findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/chkSUID_ST_NODE_DEFAULTS-SPDB").Selected = True     'Print immed.
    End With
End Sub

'not in use
'Options -> Advaned -> General -> Ignore other applications that use Dynamic Dynamic Data Exchange (DDE)
Sub IgnoreOtherApplications(ByVal Ignore As Boolean)
    Application.IgnoreRemoteRequests = Ignore
    
    On Error Resume Next
        AddIns("Analysis ToolPak").Installed = False
        AddIns("Analysis ToolPak - VBA").Installed = False
        AddIns("Euro Currency Tools").Installed = False
        AddIns("Solver Add-in").Installed = False
    On Error GoTo 0
End Sub


Sub CloseWb(ByVal wbName As String)
    Application.Wait Now + TimeValue("00:00:05")
    Dim wb As Workbook
    'close the downloaded wb
    For Each wb In Workbooks
        If wb.Name Like "*" & wbName & "*" Then
            wb.Close SaveChanges:=False
        End If
    Next
End Sub

'FunctionName  : TextToNumber
''''''Purpose  : Text store as a number changed to number
''''''Example  : colName = A, destRng = Range("A1")
Function TextToNumber(colName As String, destRng As Range)
    Columns(colName).TextToColumns destination:=destRng, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                                 Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
End Function


'FunctionName  : TextToData
''''''Purpose  : date store as a text changed to date
''''''Example  : colName = A, destRng = Range("A1")
Function TextToDate(colName As String, destRng As Range)
    Columns(colName).TextToColumns destination:=destRng, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                                   Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
End Function


'FunctionName  : NumberToText
''''''Purpose  : Number To Text Format
Function NumberToText(rng As Range)
        rng.TextToColumns _
        destination:=rng, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        FieldInfo:=Array(1, 2), _
        TrailingMinusNumbers:=True
End Function


''''''SubName  : textToColumns
''''Ex of use  : Call textToColumns(ws_FBL5N_output.Range("A1:A3000"), "|", ws_FBL5N_output.Range("A1"))
Sub TextToColumns(rng_input As Range, ByVal delimiter As String, rng_dest As Range)
    rng_input.TextToColumns destination:=rng_dest, _
                            DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, _
                            ConsecutiveDelimiter:=False, _
                            Tab:=False, _
                            Semicolon:=False, _
                            Comma:=False, _
                            Space:=False, _
                            Other:=True, _
                            OtherChar:=delimiter
End Sub



'FunctionName  : findValueRow
''''''Purpose  : Find the row with a value. Dest - where to look, lookingVal - what to find
Function findValueRow(dest As Range, ByVal lookingVal As String) As Long
    Dim findValue As Range
    Set findValue = dest.Find(What:=lookingVal, _
                                LookIn:=xlFormulas2, _
                                LookAt:=xlPart, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False, _
                                SearchFormat:=False)
    findValueRow = findValue.Row
End Function


''''''SubName     : CheckSheet
''''''Purpose     : Check is ws exists, if not, add it
Sub CheckSheet(wsName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(wsName)
    If Err.Number = 9 Then
        Set ws = Worksheets.Add(After:=Sheets(Worksheets.Count))
        ws.Name = wsName
    End If
    On Error GoTo 0
End Sub



''''''SubName     : unprotectWs
Sub unprotectWs(wsName As Worksheet, password As String)
    wsName.Unprotect password
End Sub


''''''SubName     : protectWs
Sub protectWs(wsName As Worksheet, password As String)
     wsName.Protect password:=password, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub


''''''SubName  : Paste_from_Clipboard
''''''Purpose  : Paste from cliboard
Sub Paste_from_Clipboard(dest As Range)
    Dim clipboard As Object 'Late-bound object
    Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.GetFromClipboard
    
    If clipboard.GetFormat(1) Then dest.PasteSpecial xlPasteAll 'if indicates text format
End Sub


''''''SubName  : ClearClipboard
''''''Purpose  : Find the row with a value. Dest - where to look, lookingVal - what to find
Sub ClearClipboard()
    Dim clipboard As Object 'Late-bound object
    Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.SetText ""
    clipboard.PutInClipboard
End Sub
 
 

 
 
''''''Subname  : StartSpeeding
''''''Purpose  : Turning off some functionalities to speed up the code
Sub StartSpeeding()
With Application
    .DisplayStatusBar = False
    .DisplayAlerts = False
    .EnableEvents = False
    .AskToUpdateLinks = False
    .ScreenUpdating = False
End With
End Sub



''''''Subname  : StartSpeeding
''''''Purpose  : Turning on functionalities
Sub StopSpeeding()
With Application
    .DisplayStatusBar = True
    .DisplayAlerts = True
    .EnableEvents = True
    .AskToUpdateLinks = True
    .ScreenUpdating = True
    .CutCopyMode = False
End With
End Sub



