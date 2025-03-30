Dim objArgs, folderPath
Dim objExcel, objFSO, objFolder, objSubFolder, objFile, objWorkbook
Dim objLogFile

' Get the folder path from command-line argument
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    LogMessage "Error: No folder path provided."
    WScript.Quit
End If
folderPath = objArgs(0) ' Use the provided folder path dynamically

' Create Excel application object
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False ' Keep Excel hidden during the process

' Create File System Object to access files and folders
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Open a log file for writing messages
Set objLogFile = objFSO.OpenTextFile(folderPath & "\conversion_log.txt", 8, True)

' Function to log messages instead of printing to console
Sub LogMessage(msg)
    objLogFile.WriteLine(Now & " - " & msg)
End Sub

' Check if the folder exists
If Not objFSO.FolderExists(folderPath) Then
    LogMessage "Folder not found: " & folderPath
    WScript.Quit
End If

' Get the main folder
Set objFolder = objFSO.GetFolder(folderPath)

' Loop through each subfolder in the main folder
For Each objSubFolder In objFolder.SubFolders
    ' Loop through each file in the subfolder
    For Each objFile In objSubFolder.Files
        ' Check if the file has an .xls or .xlsx extension
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "xls" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "xlsx" Then
            ' Open the Excel file
            On Error Resume Next ' Ignore errors
            Set objWorkbook = objExcel.Workbooks.Open(objFile.Path)
            If Err.Number <> 0 Then
                LogMessage "Error opening file: " & objFile.Path & " - " & Err.Description
                Err.Clear
            Else
                On Error GoTo 0 ' Restore normal error handling
                
                ' Loop through each sheet and apply the page layout changes
                Dim objSheet
                For Each objSheet In objWorkbook.Sheets
                    On Error Resume Next ' Ignore errors inside the sheet processing block
                    
                    ' Remove footers (clear left, center, and right footers)
                    With objSheet.PageSetup
                        .CenterFooter = ""
                        .LeftFooter = ""
                        .RightFooter = ""
                    End With
                    
                    ' Apply page setup to scale to fit width and height
                    With objSheet.PageSetup
                        .Zoom = False ' Disable Zoom
                        .FitToPagesWide = 1 ' Fit to 1 page wide
                        .FitToPagesTall = 1 ' Let the height flow naturally
                        .PrintArea = "" ' Clear the print area
                    End With
                    If Err.Number <> 0 Then
                        LogMessage "Error on sheet: " & objSheet.Name & " - " & Err.Description
                        Err.Clear
                    End If

                    ' Autofit columns with numeric data and increase width
                    Dim lastCol, col, lastRow, cell
                    lastCol = objSheet.Cells(1, objSheet.Columns.Count).End(-4159).Column ' Get last used column
                    lastRow = objSheet.Cells(objSheet.Rows.Count, 1).End(-4162).Row ' Get last used row

                    ' Loop through each column from D (column 4) to last column
                    For col = 4 To lastCol
                        Dim containsNumber
                        containsNumber = False
                        For row = 1 To lastRow
                            Set cell = objSheet.Cells(row, col)
                            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                                containsNumber = True
                                Exit For
                            End If
                        Next
                        ' If the column contains numbers, autofit and increase width
                        If containsNumber Then
                            objSheet.Columns(col).AutoFit
                            objSheet.Columns(col).ColumnWidth = objSheet.Columns(col).ColumnWidth + 5
                        End If
                    Next
                    On Error GoTo 0 ' Restore normal error handling
                Next
                
                ' Enforce the output PDF file path explicitly using the parent folder
                On Error Resume Next
                Dim outputPDF
                outputPDF = objSubFolder.Path & "\" & objFSO.GetBaseName(objFile.Name) & ".pdf"
                objWorkbook.ExportAsFixedFormat 0, outputPDF
                If Err.Number <> 0 Then
                    LogMessage "Error exporting to PDF for file: " & objFile.Path & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' Close the workbook without saving
                objWorkbook.Close False
            End If
        End If
    Next
Next

' Quit Excel application
objExcel.Quit

' Clean up objects
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
objLogFile.Close
Set objLogFile = Nothing
