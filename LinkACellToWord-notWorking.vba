Sub LinkExcelTable()
    Dim objExcel As Object
    Dim wb As Workbook
    Dim strPath As String
    Dim cellExists As Boolean
    cellExists = False

    ' Update the path to your Excel workbook, including the filename and extension
    strPath = "C:\\Kiingo\\Assurance Excel2Doc Macro\\generated_workbook_1.xlsx"

    ' Create an Excel object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False ' Ensure Excel runs in the background

    ' Attempt to open the Excel file
    On Error Resume Next ' Disable error interruption
    Set wb = objExcel.Workbooks.Open(strPath)
    If Err.Number <> 0 Then
        MsgBox "No workbook opened"
        Err.Clear
        objExcel.Quit
        Set objExcel = Nothing
        Exit Sub
    End If
    On Error GoTo 0 ' Turn back on normal error handling

    ' Check for the existence of the cell in the specified sheet and range
    On Error Resume Next
    Dim cellValue As Variant
    cellValue = wb.Sheets("Data").Range("A1").Value
    If Err.Number = 0 Then
        cellExists = True
    Else
        MsgBox "Cell does not exist"
        Err.Clear
    End If
    On Error GoTo 0

    ' Add the Excel worksheet object to the Word document if the cell exists
    If cellExists Then
        With ActiveDocument
            .Shapes.AddOLEObject ClassType:="Excel.Sheet", _
                FileName:=strPath, Link:=True, _
                DisplayAsIcon:=False, Range:="Data!A1"
        End With
        MsgBox "Linked cell A1 from 'Data' sheet."
    Else
        MsgBox "The specified cell 'A1' does not exist in the Excel file."
    End If

    ' Close the workbook
    wb.Close SaveChanges:=False

    ' Quit Excel
    objExcel.Quit
    Set objExcel = Nothing
End Sub
