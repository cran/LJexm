Dim objArgs, folderPath
Dim wordApp, objFSO, objFolder, objSubFolder, objFile, wordDoc
Dim objLogFile

' Get the folder path from command-line argument
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    LogMessage "Error: No folder path provided."
    WScript.Quit
End If
folderPath = objArgs(0) ' Use the provided folder path dynamically

' Create a new instance of Word Application
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False ' Keep Word hidden during the process

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
        ' Check if the file has a .doc or .docx extension
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "doc" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "docx" Then
            On Error Resume Next ' Start error handling
            ' Open the Word document
            Set wordDoc = wordApp.Documents.Open(objFile.Path)
            
            If Err.Number <> 0 Then
                LogMessage "Error opening file: " & objFile.Path & " - " & Err.Description
                Err.Clear
                On Error GoTo 0
                ' Skip this file by jumping to the next file in the loop
            Else
                ' Enforce the output PDF file path explicitly using the parent folder of the .doc file
                Dim pdfPath
                pdfPath = objSubFolder.Path & "\" & objFSO.GetBaseName(objFile.Name) & ".pdf"
                ' Save the document as PDF
                wordDoc.SaveAs2 pdfPath, 17 ' 17 is the FileFormat for PDF
                
                ' Check if there was an error during the save process
                If Err.Number <> 0 Then
                    LogMessage "Error saving file as PDF: " & objFile.Path & " - " & Err.Description
                    Err.Clear
                End If
                ' Close the document
                wordDoc.Close False
            End If
            On Error GoTo 0 ' Restore normal error handling
        End If
    Next
Next

' Quit Word application
wordApp.Quit

' Clean up objects
Set wordDoc = Nothing
Set wordApp = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
objLogFile.Close
Set objLogFile = Nothing