'-------------------------------------------------------------------------------------------------------------------
'This will scan a selected drive, tell you the names of the files in it, and give you the dates they were created.
'To run, place anywhere and double click. Output saved to Desktop\Scan\Results.txt
'-------------------------------------------------------------------------------------------------------------------

driveLetter = InputBox("Enter the drive letter to scan (e.g. C:):") 							'Prompts user for the drive letter

If Right(driveLetter, 1) <> ":" Then 															'Ensures the drive letter is formatted correctly
    driveLetter = driveLetter & ":"
End If

fn = driveLetter & "\" 																			'Defines the folder to check

Set fs = CreateObject("Scripting.FileSystemObject") 											'Creates FileSystemObject

If fs.DriveExists(driveLetter) Then 															'Checks if the drive exists
    msgbox "Good drive" 																		'Shows "Good drive" if drive exists
Else
    msgbox "Bad drive" 																			'Shows "Bad drive" if drive doesn't exist
    WScript.Quit 																				'Exits if the drive does not exist
End If

Set shell = CreateObject("WScript.Shell") 														'Creates WScript.Shell
outputFilePath = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\Scan\Results.txt" 	'Defines the output file path on the user's desktop


Dim fileIndex																					'Checks if output file already exists and creates a new filename if it does
fileIndex = 1
While fs.FileExists(outputFilePath)
    outputFilePath = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\Scan\Results" & fileIndex & ".txt"
    fileIndex = fileIndex + 1
Wend

Set outputFile = fs.CreateTextFile(outputFilePath, True)										'Opens the output file for writing

Sub ScanFolder(folder)																			'Function to recursively scan folders and write file details to output	
    Dim file, subfolder
    For Each file In folder.Files																'Loops through each file in the folder
        outputFile.WriteLine(file.Path & " - Created: " & file.DateCreated)
    Next
    For Each subfolder In folder.SubFolders 													'Loops through each subfolder in the folder
        ScanFolder subfolder 																	'Recursive call to scan subfolder
    Next
End Sub

On Error Resume Next 																			'Enable error handling
Set rootFolder = fs.GetFolder(fn) 																'Starts scanning from the root folder
If Err.Number <> 0 Then
    outputFile.WriteLine "Error accessing folder: " & Err.Description & " (Error Number: " & Err.Number & ")"
    MsgBox "Error accessing folder: " & Err.Description
    WScript.Quit
End If
On Error GoTo 0 																				'Disables error handling

ScanFolder rootFolder

outputFile.Close 																				'Closes the output file

msgbox "Results File Created and output saved to " & outputFilePath 							'Shows message box indicating Results.txt has been created
