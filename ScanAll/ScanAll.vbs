'-------------------------------------------------------------------------------------------------------------------
'This will scan all available drives, tell you the names of the files in them, and give you the dates they were created.
'To run, place anywhere and double click. Output saved to the current folder as ResultsAll.txt
'-------------------------------------------------------------------------------------------------------------------

Set fs = CreateObject("Scripting.FileSystemObject")  																	' Creates FileSystemObject
Set shell = CreateObject("WScript.Shell")            																	' Creates WScript.Shell

outputFilePath = fs.GetParentFolderName(WScript.ScriptFullName) & "\ResultsAll.txt"										' Define the output file path in the same directory as the script

Dim fileIndex																											' Check if the output file already exists and create a new filename if it does
fileIndex = 1
While fs.FileExists(outputFilePath)
    outputFilePath = fs.GetParentFolderName(WScript.ScriptFullName) & "\ResultsAll" & fileIndex & ".txt"
    fileIndex = fileIndex + 1
Wend

Set outputFile = fs.OpenTextFile(outputFilePath, 8, True)  																' Open the output file for appending 8 = ForAppending

Sub ScanFolder(folder)																									' Subroutine to scan a folder and write file details to output
    Dim file, subfolder
    For Each file In folder.Files  																						' Loops through each file in the folder
        outputFile.WriteLine(file.Path & " - Created: " & file.DateCreated)
    Next
    For Each subfolder In folder.SubFolders  																			' Loops through each subfolder in the folder
        ScanFolder subfolder  																							' Recursive call to scan subfolder
    Next
End Sub

Dim driveLetter																											' Loop through all drives from A to Z
For driveLetter = Asc("A") To Asc("Z")
    Dim drivePath
    drivePath = Chr(driveLetter) & ":"  																				' Convert ASCII value back to character

    If fs.DriveExists(drivePath) Then  																					' Checks if the drive exists
        On Error Resume Next  																							' Enable error handling
        Set rootFolder = fs.GetFolder(drivePath)  																		' Starts scanning from the root folder
        If Err.Number <> 0 Then
            outputFile.WriteLine "Error accessing folder: " & Err.Description & " (Error Number: " & Err.Number & ")"
        Else
            ScanFolder rootFolder  																						' Scan the valid drive
        End If
        On Error GoTo 0  																								' Disables error handling
    End If
Next

outputFile.Close  																										' Closes the output file

MsgBox "Results File Created and output saved to " & outputFilePath  													' Shows message box indicating ResultsAll.txt has been created
