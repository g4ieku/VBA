Attribute VB_Name = "mod_SendEmail_Helper"
Option Explicit
' Module description: Send a basic email using Outlook


''''''''SubName:   SendEmail (Takes parameters to send emails based on the needed things)
'Example of use:  Call SendEmail("Who@gmail.com", "subject", "body", includeSignature:=True)

Public Sub SendEmail(ByVal emailTo As String _
                    , ByVal subject As String _
                    , ByVal body As String _
                    , Optional ByVal CC As String = "" _
                    , Optional ByVal BCC As String = "" _
                    , Optional ByVal includeSignature As Boolean = False _
                    , Optional ByVal recipientsRange As Range = Nothing _
                    , Optional ByVal recipientsArray As Variant = Empty _
                    , Optional ByVal attachmentsRange As Range = Nothing _
                    , Optional ByVal attachmentsArray As Variant = Empty _
                    , Optional ByVal attachmentsString As String = "")
    
    ' Tools->Reference and check the box beside "Microsoft Outlook 16.0 Object Library" 'early biding
    'Dim myOutlook As New Outlook.Application
    
    ' Create email
    Dim outApp As Object: Set outApp = CreateObject("Outlook.Application")
    Dim outMail As Object: Set outMail = outApp.CreateItem(0)
        
    ' Configure email
    With outMail
    
        .To = emailTo
        .subject = subject
        .CC = CC
        .BCC = BCC
        
        If includeSignature = True Then
            .Display
            If Len(Trim(.htmlBody)) <> 0 Then
                .htmlBody = body & "<br><br>" & .htmlBody
            End If
        Else
            .body = body
        End If

        If Not recipientsRange Is Nothing Then
            Call RangeToCollection(recipientsRange, .recipients)
        ElseIf IsEmpty(recipientsArray) = False Then
            Call ArrayToCollection(recipientsArray, .recipients)
        End If

        If Not attachmentsRange Is Nothing Then
            Call RangeToCollection(attachmentsRange, .attachments)
        ElseIf IsEmpty(attachmentsArray) = False Then
            Call ArrayToCollection(attachmentsArray, .attachments)
        ElseIf Len(Trim(attachmentsString)) > 0 Then
            Call StringToCollection(attachmentsString, .attachments)
        End If
        
        .Display ' Turn on for testing
  '      .Send
    End With

End Sub

Private Sub RangeToCollection(rg As Range, coll As Object)
    With rg
        Dim files As Variant
        If .Cells.Count = 1 Then
            ReDim files(1 To 1, 1 To 1)
            files(1, 1) = rg.Value
        Else
            If .Rows.Count > 1 Then
                files = rg.Value
            ElseIf .Columns.Count > 1 Then
                files = WorksheetFunction.Transpose(.Value)
            End If
        End If
    End With
    Dim i As Long
    For i = LBound(files) To UBound(files)
        If IsEmpty(files) Then GoTo continue
        coll.Add files(i, 1)
continue:
    Next i
End Sub

Private Sub ArrayToCollection(items As Variant, coll As Object)
    Dim item As Variant
    For Each item In items
        coll.Add item
    Next item
End Sub

Private Sub StringToCollection(items As String, coll As Object)
    Dim item As Variant
    For Each item In Split(items, ",")
        coll.Add item
    Next item
End Sub


Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    
    Call StartSpeeding
    
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    
    Set TempWB = Workbooks.Add(1)
    
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
            .DrawingObjects.Visible = True
            .DrawingObjects.Delete
        On Error GoTo 0
    End With
    
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        FileName:=TempFile, _
        sheet:=TempWB.Sheets(1).Name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", "align=left x:publishsource=")
    
    TempWB.Close SaveChanges:=False
    Kill TempFile
    
    Call StopSpeeding
    
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


