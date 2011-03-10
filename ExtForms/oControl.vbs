dim oControl
Dim objFileSystem, objOutputFile
Dim strOutputFile

set oControl = WScript.CreateObject("SDEControl.SDEControl")

oControl.Domain="bbtspb.ru"
oControl.SupportID="2f03a7c7-efe5-4373-9108-f954b866e67b"
oControl.GrantedName = "Baltiysky Balkerniy Terminal"
oControl.Count=25	
oControl.AdditionalCount = 10 
oControl.ProductID = 45

' generate a filename base on the script name
strOutputFile = "./" & Split(WScript.ScriptName, ".")(0) & ".xml"

Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objOutputFile = objFileSystem.CreateTextFile(strOutputFile, TRUE)

objOutputFile.WriteLine(oControl.GetLicense())
objOutputFile.Close

Set objFileSystem = Nothing

WScript.Quit(0)


