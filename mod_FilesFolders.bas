Attribute VB_Name = "mod_FilesFolders"
Option Explicit

Public fso As Object
Public FileName As String
Public Const MainFolder As String = "C:\Reports\"

'Shell "explorer.exe " & SubFolderDownlReports, vbNormalFocus                      'display folder

'FunctionName  : FolderCreate
''''''Purpose  : Creating Folder
Function FolderCreate(ByVal Path As String) As Boolean
    FolderCreate = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If FolderExists(Path) Then
        FolderCreate = False
        Exit Function
    Else
        fso.CreateFolder Path
    End If
End Function



'FunctionName  : FolderExists
''''''Purpose  : Checking if folder exist
Function FolderExists(ByVal Path As String) As Boolean
    FolderExists = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(Path) Then FolderExists = True
End Function




'FunctionName  : filePath
''''''Purpose  : Taking file path from the msoFileDialogFilePicker
Function FilePath(ByVal TitleText As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
            .Title = TitleText
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel Files Only", "*.xls*"
            '.Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm"
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        Else: Exit Function
        End If
    End With
End Function



'FunctionName  : GetWorkbook
''''''Purpose  : to check if the workbook is currently open. If not, then it will open the workbook. In either case you will end up with the workbook opened.
Function GetWorkbook(ByVal sFullFilename As String) As Workbook
    Dim sFilename As String
    sFilename = Dir(sFullFilename)
    
    On Error Resume Next
    Dim wk As Workbook
    Set wk = Workbooks(sFilename)
    
    If wk Is Nothing Then
        Set wk = Workbooks.Open(sFullFilename)
    End If
    
    On Error GoTo 0
    Set GetWorkbook = wk
End Function

Sub ExampleOpenWorkbook()
    Dim sFilename As String
    sFilename = "C:\Docs\Book2.xlsx"

    Dim wk As Workbook
    Set wk = GetWorkbook(sFilename)
End Sub



'FunctionName  : IsWorkBookOpen
''''''Purpose  : Function to check if workbook is already open
Function IsWorkBookOpen(strBookName As String) As Boolean
    Dim oBk As Workbook
    
    On Error Resume Next
    Set oBk = Workbooks(strBookName)
    On Error GoTo 0
    
    If Not oBk Is Nothing Then
        IsWorkBookOpen = True
    End If
End Function

Sub ExampleUse()
    Dim sFilename As String
    sFilename = "C:\temp\writedata.xlsx"

    If IsWorkBookOpen(Dir(sFilename)) = True Then
        MsgBox "File is already open. Please close file and run macro again."
        Exit Sub
    End If
    
    ' Write to workbook here
End Sub



'FunctionName  : UserSelectWorkbook
''''''Purpose  : The function returns the full file name if a file was selected. If the user cancels it displays a message and returns an empty string.
Public Function UserSelectWorkbook() As String
    On Error GoTo ErrorHandler

    Dim sWorkbookName As String
    Dim FD As FileDialog: Set FD = Application.FileDialog(msoFileDialogFilePicker)

    With FD                                                 ' Open the file dialog
        .Title = "Please Select File"                       ' Set Dialog Title
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm"   ' Add filter
        .AllowMultiSelect = False                           ' Allow selection of one file only
        .Show                                               ' Display dialog

        If .SelectedItems.Count > 0 Then
            UserSelectWorkbook = .SelectedItems(1)
        Else
            MsgBox "Selecting a file has been cancelled. "
            UserSelectWorkbook = ""
        End If
    End With

    Set FD = Nothing                                        ' Clean up
Done:
    Exit Function
ErrorHandler:
    MsgBox "Error: " + Err.Description
End Function



Public Sub TestUserSelect()
    Dim userBook As Workbook, sFilename As String

    sFilename = UserSelectWorkbook()                            ' Call the UserSelectworkbook function

    If sFilename <> "" Then                                     ' If the filename returns is blank the user cancelled
        Set userBook = Workbooks.Open(sFilename)                ' Open workbook and do something with it
    End If
End Sub


'workbook cannot be store on the desktop - SQL reason )
Function checkIfFileOnDesktop() As Boolean
    checkIfFileOnDesktop = False
    
    If InStr(1, ThisWorkbook.FullName, "https://heiway-my.sharepoint.com") >= 1 Then
        MsgBox Prompt:="You cannot run the code in the file stored on your desktop. Move your file to a different location and run once again!", Buttons:=vbOKOnly + vbInformation, Title:="File located on your desktop"
        checkIfFileOnDesktop = True
    End If
End Function
