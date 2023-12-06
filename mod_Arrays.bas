Attribute VB_Name = "mod_Arrays"
Option Explicit



'''''Sub  : CopyRangeToArray
Public Sub CopyRangeToArray(ByVal FileName As String _
                        , ByVal sheetName As String _
                        , ByRef data As Variant)
    Dim book As Workbook
    Set book = Workbooks.Open(FileName, ReadOnly:=True)
    
    Dim sheet As Worksheet
    Set sheet = book.Worksheets(sheetName)
    
    data = sheet.Range("A1").CurrentRegion
    
    book.Close SaveChanges:=False
End Sub

'FunctionName  : IsInArray
''''''Purpose  : Checking if an item is inside the array
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If CStr(arr(i)) = CStr(stringToBeFound) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


'''''Sub  : ArrayToRange
Public Sub ArrayToRange(ByRef arr As Variant, rg As Range)
    rg.Resize(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1) = arr
End Sub



'FunctionName  : ArrayLength
''''''Purpose  :  to return the number of items in any array no matter how many dimensions
Function ArrayLength(arr As Variant) As Long
    On Error GoTo eh
    
    ' Loop is used for multidimensional arrays. The Loop will terminate when a
    ' "Subscript out of Range" error occurs i.e. there are no more dimensions.
    Dim i As Long, length As Long
    length = 1
    
    ' Loop until no more dimensions
    Do While True
        i = i + 1
        ' If the array has no items then this line will throw an error
        length = length * (UBound(arr, i) - LBound(arr, i) + 1)
        ' Set ArrayLength here to avoid returing 1 for an empty array
        ArrayLength = length
    Loop

Done:
    Exit Function
eh:
    If Err.Number = 13 Then ' Type Mismatch Error
        Err.Raise vbObjectError, "ArrayLength" _
            , "The argument passed to the ArrayLength function is not an array."
    End If
End Function



''''''SubName  : QuickSort
''''''Purpose  : can be used to sort an array
Sub QuickSort(arr As Variant, first As Long, last As Long)
  Dim vCentreVal As Variant, vTemp As Variant
  
  Dim lTempLow As Long
  Dim lTempHi As Long
  lTempLow = first
  lTempHi = last
  
  vCentreVal = arr((first + last) \ 2)
  Do While lTempLow <= lTempHi
  
    Do While arr(lTempLow) < vCentreVal And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    
    Do While vCentreVal < arr(lTempHi) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    
    If lTempLow <= lTempHi Then
    
        vTemp = arr(lTempLow)               ' Swap values

        arr(lTempLow) = arr(lTempHi)
        arr(lTempHi) = vTemp
      
        ' Move to next positions
        lTempLow = lTempLow + 1
        lTempHi = lTempHi - 1
      
    End If
    
  Loop
  
  If first < lTempHi Then QuickSort arr, first, lTempHi
  If lTempLow < last Then QuickSort arr, lTempLow, last
End Sub

'example of use
Sub TestSort()
    ' Create temp array
    Dim arr() As Variant
    arr = Array("Banana", "Melon", "Peach", "Plum", "Apple")
  
    ' Sort array
    QuickSort arr, LBound(arr), UBound(arr)

    ' Print arr to Immediate Window(Ctrl + G)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next i
End Sub


''''''Example of use
'Dim arr_Output As Variant: arr_Output = RecordsetToArray(SQLQuery, ThisWorkbook.FullName)  'SQL Query write to array, later we will transpose the array from columns|rows format to rows | columns
'Dim arr_OutputTransp As Variant
'Call TransposeArray(arr_Output, arr_OutputTransp)                                          'to have standard format for array: rows|columns

''''''FunctionName  : RecordsetToArray
'''''''''''Purpose  : Write range to the array
''''''''Important!  : The array doesn't look in the standard way when we use sql query, use later the function 'TransposeArray'
'                     to have standard format for array: rows|columns. Check the example above!
Function RecordsetToArray(SQLquery As String, DataFilePath As String) As Variant
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object

    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                            "Data Source=" & DataFilePath & ";" & _
                            "Extended Properties='Excel 12.0 Xml;HDR=YES';"
    conn.Open
    Set rs = CreateObject("ADODB.Recordset")
    rs.ActiveConnection = conn
    rs.Source = SQLquery
    rs.Open
    
    'GetRows - return an array from the record set. IT RETURNS ARRAY IN THE COLUMNS|ROWS FORMAT, not in the rows|columns format!
    'transpose columns to rows by using sub: 'TransposeArray' to have a standard format for array
    RecordsetToArray = rs.GetRows
    'rs.MoveFirst                                             'unblock that if you would like to write to the worksheet
    'ThisWorkbook.Worksheets("Return POs").Range("A1").CopyFromRecordset rs 'write to worksheet - for debugging

    rs.Close
    conn.Close
End Function

''''''FunctionName  : TransposeArray
'''''''''''Purpose  : Transposing array, useful when you read range to array by using sql query and later you
'                     would like to have the array in the standart format - rows|columns. Check the example above!
Public Sub TransposeArray(ByRef InputArr As Variant, ByRef ReturnArray As Variant)
    Dim rowIndex As Long, colIndex As Long
    Dim LBound1 As Long, LBound2 As Long, UBound1 As Long, UBound2 As Long

    LBound1 = LBound(InputArr, 1)
    LBound2 = LBound(InputArr, 2)
    UBound1 = UBound(InputArr, 1)
    UBound2 = UBound(InputArr, 2)

    ReDim ReturnArray(LBound2 To UBound2, LBound1 To UBound1)

    For rowIndex = LBound2 To UBound2
    For colIndex = LBound1 To UBound1
        ReturnArray(rowIndex, colIndex) = InputArr(colIndex, rowIndex)
    Next colIndex, rowIndex
End Sub


'FunctionName: ReturnOneColumnFromArray
'Purpose     : Returning one column from array with multiple columns (columnToExtract - which one should be returned), outRange - where to output the array
Sub ReturnOneColumnFromArray(orginalArray As Variant, columnToExtract As Integer, outRange As Range, Optional formatDate As Boolean = False)
    Dim columnArray() As Variant
    Dim numRows As Long: numRows = UBound(orginalArray, 1)              'Get the number of rows in the original array
    ReDim columnArray(0 To numRows, 1 To 1)                             'Initialize the columnArray with the same number of rows

    ' Loop through the original array and extract the specified column
    For i = 0 To numRows
        If formatDate = False Then                                      'Arrays automaticly change date format from dd/mm/yyyy, we would like to 'reverse' this process
            columnArray(i, 1) = orginalArray(i, columnToExtract)
        Else: columnArray(i, 1) = Format(orginalArray(i, columnToExtract), "mm/dd/yyyy")
        End If
    Next i

    outRange = columnArray                                              'write data to the worksheet range
End Sub

'SubName     : WriteArrayToRange
'Purpose     : Writing an array to range
Sub WriteArrayToRange(arr As Variant, colNr As Long)
    Report.Cells(2, colNr).Resize(UBound(arr, 1) + 1, UBound(arr, 2)).Value = arr
End Sub
