Sub SplitDataIntoFiles()
    Dim wsData As Worksheet
    Dim uniqueCountries As Collection
    Dim lastRow As Long
    Dim cell As Range
    Dim country As String
    Dim folderPath As String

    Set wsData = ThisWorkbook.Worksheets("split table")
    lastRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    folderPath = "C:\Users\" & Environ("USERNAME") & "\Desktop\testTask\Country\"
    Set uniqueCountries = New Collection

    ' Collect unique countries to collection
    For Each cell In wsData.Range("C2:C" & lastRow)
        country = cell.Value
        On Error Resume Next
        uniqueCountries.Add country, CStr(country)
        On Error GoTo 0
    Next cell

    ' Create the main folder path
    createFolderIfNotExist ("C:\Users\" & Environ("USERNAME") & "\Desktop\testTask\")
    createFolderIfNotExist (folderPath)

    ' Loop through unique countries and create separate files
    Dim countryName As Variant
    
    For Each countryName In uniqueCountries
        ' Filter and copy data for the current country
        wsData.AutoFilterMode = False
        wsData.Range("C1").AutoFilter Field:=3, Criteria1:=countryName
        wsData.UsedRange.SpecialCells(xlCellTypeVisible).Copy

        ' Create and save new Workbook
        Dim newWb As Workbook
        Set newWb = Workbooks.Add
        newWb.Sheets(1).Range("A1").PasteSpecial
        newWb.SaveAs folderPath & countryName & ".xlsx"
        newWb.Close SaveChanges:=False
    Next countryName
    
    wsData.AutoFilterMode = False

    MsgBox "Data has been split"
End Sub

Function createFolderIfNotExist(path As String)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
End Function
