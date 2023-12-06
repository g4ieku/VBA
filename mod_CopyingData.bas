Attribute VB_Name = "mod_CopyingData"
Option Explicit

''''Subname  : CopyDataAssign
' Description: Copy the data using the assignment. This is much faster than Range copy
'               if you have a lot of data and are doing multiple copies.
Sub CopyDataAssign()
    Dim FileName As Variant                                                 ' Ask the user to select an Excel workbook
    FileName = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx")
    
    If FileName <> False Then

        Dim wksource As Workbook                                            ' Get the workbook from the filename
        Set wksource = Workbooks.Open(FileName, ReadOnly:=True)
        
        Dim rgSource As Range                                               ' Get the range from the worksheet
        Set rgSource = wksource.Worksheets(1).Range("A1").CurrentRegion
        
        Dim rgDest As Range                                                 ' Copy using the assignment method which is faster than range copy
        Set rgDest = ThisWorkbook.Worksheets("Data").Range("A1")
        rgDest.Resize(rgSource.Rows.Count, rgSource.Columns.Count).Value = rgSource.Value
           
        wksource.Close SaveChanges:=False                                   ' Close the file
    
    End If
End Sub



''''Subname  : CopyDataBest
' Description: This is the most optimised code. It uses arrays to store the data.
'               Use two functions to make the code neater and readable while maintainn the performance.
Sub CopyDataBest()
    Dim FileName As Variant                                                 ' Ask the user to select an Excel workbook
    FileName = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx")
    
    If FileName <> False Then
        
        Dim data As Variant                                                 ' Copy the range to an array
        Call CopyRangeToArray(FileName, "Fruit", data)
        
        Dim i As Long                                                       ' Process data here
        For i = LBound(data, 1) To UBound(data, 1)
            ' update date here e.g. change all fruit to Mango. data(i,2)= "Mango"
        Next i
           
        Call ArrayToRange(data, shData.Range("A1"))                         ' Write the array to the range
    
    End If
End Sub



''''Subname  : CopyFromMultipleFiles
' Description: This is the most optimised code. Use two functions to make the code
'               neater and readable while maintainn the performance.
Sub CopyFromMultipleFiles()
    Dim FileName As Variant                                             ' Ask the user to select one or more Excel workbooks
    FileName = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", MultiSelect:=True)
    
    If IsArray(FileName) = True Then
        
        shData.Cells.ClearContents                                       ' Clear existing data
        
        Dim file As Variant, data As Variant, rgDest As Range
        For Each file In FileName
            
            Call CopyRangeToArray(file, "Fruit", data)                   ' Copy the range to an array
                       
            Set rgDest = shData.Range("A" & shData.Rows.Count).End(xlUp) ' Get the last row with data
            If rgDest.Value <> "" Then Set rgDest = rgDest.Offset(1)
            
            Call ArrayToRange(data, rgDest)                              ' Write the array to the range
        
        Next file
    
    End If
End Sub

