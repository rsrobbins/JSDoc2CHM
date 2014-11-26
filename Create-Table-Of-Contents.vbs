REM *****************************************************************
REM * VBScript To Convert index.html to HxT table of contents file	*
REM * by Robert S. Robbins                                          *
REM * written on 11/25/2014                                         *
REM *****************************************************************
REM
CONST adSaveCreateOverWrite = 2
REM Create file system objects 
Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
strWorkingDirectory = Left(WScript.ScriptFullName,  InStrRev(WScript.ScriptFullName,"\")-1)

Set input_file = objFileSystem.OpenTextFile(strWorkingDirectory & "/index.html") 
Set output_file = objFileSystem.CreateTextFile (strWorkingDirectory & "/Famous.HxT", True, False)

REM Create a regular expression object
Set rxLinkText = new RegExp
rxLinkText.Pattern = """>(.*?)</a>"
rxLinkText.Global = True
rxLinkText.IgnoreCase = True

Set rxUrl = new RegExp
rxUrl.Pattern = "href=""(.*?)"""
rxUrl.Global = True
rxUrl.IgnoreCase = True

' Write the initial lines needed for a HxT file
output_file.WriteLine "<?xml version=""1.0""?>"
output_file.WriteLine "<!DOCTYPE HelpTOC SYSTEM ""ms-help://hx/resources/HelpTOC.DTD"">"
output_file.WriteLine "<HelpTOC"
output_file.WriteLine "   DTDVersion=""1.0"""
output_file.WriteLine "   FileVersion=""1.0"""
output_file.WriteLine "   LangId=""1033"">"

' Create flag to indicate if a XML needs to be closed
bNode = False

' Read through file
Do Until input_file.AtEndOfStream
	' Read a line. Remember, every use of input_file.Readline reads another line
	strLine = Trim(input_file.Readline)
	' Find a line to be processed
	If InStr(strLine, "<li>") > 0  Then
			If InStr(strLine, "<li><a href=") > 0  Then
				' Topic XML node
				Set oMatches = rxLinkText.Execute(strLine)
				Set oMatch = oMatches(0)
				strTitle = oMatch.SubMatches(0)
				
				Set objMatches = rxUrl.Execute(strLine)
				Set objMatch = objMatches(0)
				strUrl = objMatch.SubMatches(0)
				strUrl = Replace(strUrl, "./", "")	
				strUrl = Replace(strUrl, "/", "\")			
				output_file.WriteLine "<HelpTOCNode Title=""" & strTitle & """ Url=""" & strUrl & """/>"
		  Else
		  	' Category XML node
		  	If bNode = True Then
		  		' Close previous node
		  		output_file.WriteLine "</HelpTOCNode>"
		    End If
		  	strCategory = Replace(strLine, "<li>", "")
		  	strCategory = Replace(strCategory, "</li>", "")
		  	output_file.WriteLine "<HelpTOCNode Title=""" & strCategory & """ Icon=""0"">" 
		  	bNode = True
			End If
	End If
Loop 

' Write the final lines needed for a HxT file
output_file.WriteLine "</HelpTOCNode>"
output_file.WriteLine "   <ToolData Name=""MSTOCEXPST"" Value=""Expanded""/>"
output_file.WriteLine "   <ToolData Name=""MSTOCMRUDIR"" Value=""C:\_Win\Visual Studio Projects\WZDocumentation\""/>"
output_file.WriteLine "</HelpTOC>"


Set input_file =  Nothing
Set output_file =  Nothing

MsgBox "Done Processing!",vbInformation,"Create HxT Script"
