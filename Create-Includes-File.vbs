REM ***************************************************************
REM * VBScript To Create HxF file, the includes file            	*
REM * by Robert S. Robbins                                        *
REM * written on 11/25/2014                                       *
REM ***************************************************************
REM
CONST adSaveCreateOverWrite = 2
REM Create file system objects 
Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
strWorkingDirectory = Left(WScript.ScriptFullName,  InStrRev(WScript.ScriptFullName,"\")-1)

Set output_file = objFileSystem.CreateTextFile (strWorkingDirectory & "/Famous.HxF", True, False)

' Write the initial lines needed for a HxT file
output_file.WriteLine "<?xml version=""1.0""?>"
output_file.WriteLine "<!DOCTYPE HelpFileList SYSTEM ""ms-help://Hx/Resources/HelpFileList.DTD"">"
output_file.WriteLine "<HelpFileList DTDVersion=""1.0"">"

' Find all files in a folder
Set objFolder = objFileSystem.GetFolder(strWorkingDirectory)
Set colFiles = objFolder.Files
For Each objFile in colFiles
		strFilePath = objFile.Path
		strFilePath = Replace(strFilePath, strWorkingDirectory & "\", "")
    output_file.WriteLine "<File Url=""" & strFilePath & """/>"
Next

ShowSubfolders objFileSystem.GetFolder(objFolder)

' Write the final lines needed for a HxT file
output_file.WriteLine "</HelpFileList>"

Set input_file =  Nothing
Set output_file =  Nothing

MsgBox "Done Processing!",vbInformation,"Create HxF Script"

' Recursive function for subfolders
Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        Set objFolder = objFileSystem.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
        		strFilePath = objFile.Path
						strFilePath = Replace(strFilePath, strWorkingDirectory & "\", "")
            output_file.WriteLine "<File Url=""" & strFilePath & """/>"
        Next
        ShowSubFolders Subfolder
    Next
End Sub