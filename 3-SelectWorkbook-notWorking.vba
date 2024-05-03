Sub LinkExcelTable()
    Dim objExcel As Object
    Dim wb As Workbook
    Dim fileNameAndPath As Variant
    Dim tableExists As Boolean
    tableExists = False

    ' Update the path to your Excel workbook, including the filename and extension
    fileNameAndPath = Application.GetOpenFilename(FileFilter:="Excel Files (*.XLSX), *.XLSX", Title:="Select File To Be Opened")

    ' Create an Excel object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False ' Ensure Excel runs in the background

    ' Attempt to open the Excel file
    On Error Resume Next ' Disable error interruption
    Set wb = objExcel.Workbooks.Open(fileNameAndPath)
    On Error GoTo 0 ' Turn back on normal error handling
    
    If wb Is Nothing Then
    MsgBox "No workbook opened"
    
    End If

    ' Check if the file was opened successfully
    If Not wb Is Nothing Then
        ' Check for the existence of the table named "Table1"
        On Error Resume Next ' Disable error interruption
        Dim tbl As ListObject
        Set tbl = wb.Sheets("Data").ListObjects("Table1")
        On Error GoTo 0 ' Turn back on normal error handling

        ' Verify if the table exists
        If Not tbl Is Nothing Then
            tableExists = True
            MsgBox "table exists"
        Else
            MsgBox "Table does not exist"
        End If
        
    End If

End Sub
