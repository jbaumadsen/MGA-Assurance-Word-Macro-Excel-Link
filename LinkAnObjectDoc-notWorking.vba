Sub LinkExcelTable()
    Dim objExcel As Object
    Dim wb As Workbook
    Dim strPath As String
    Dim tableExists As Boolean
    tableExists = False

    ' Update the path to your Excel workbook, including the filename and extension
    strPath = "C:\\Kiingo\\Assurance Excel2Doc Macro\\generated_workbook_1.xlsx"

    ' Create an Excel object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False ' Ensure Excel runs in the background

    ' Attempt to open the Excel file
    On Error Resume Next ' Disable error interruption
    Set wb = objExcel.Workbooks.Open(strPath)
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

    ' Create an Excel object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False ' Ensure Excel runs in the background
        ' Add the Excel worksheet object to the Word document if the table exists
        If tableExists Then
            With ActiveDocument
                .Shapes.AddOLEObject FileName:=strPath, _
                    DisplayAsIcon:=False, IconFileName:=strPath, IconIndex:=0, _
                    IconLabel:=strPath
            End With
        Else
            MsgBox "The specified table 'Table1' does not exist in the Excel file."
        End If

        ' Close the workbook
        wb.Close SaveChanges:=False

    ' Quit Excel
    objExcel.Quit
    Set objExcel = Nothing
End Sub