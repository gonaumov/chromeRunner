'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The idea of this script is to make the starting of Google Chrome browser 
' without security much easier than it was before. The script checks if there
' is running instances of chrome.exe and if there is a present instance,
' alerts the user that the running instance must be killed before Google Crome
' is started. If the user clicks on Yes button, the script will kill all instances of 
' chrome.exe. After this, will open a new browser instance without security. 
' If the user clicks on No button, the script will do nothing. If there is no   
' present running instance of chrome.exe, the script will start 
' Google Chrome without security. How to use it? Just clone this repository 
' and fill correct path to chrome.exe. The script is tested on windows 8 
' with admin rigths. Don't hesitate to write to me with suggestions and
' feature requests. I will be happy if someone use this. :) 
' 
' @author Georgi Naumov
'  gonaumov@gmail.com for contacts and 
'  suggestions. 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Option Explicit

Dim oShell, oRet, regexNotRunning, strCommandResult, userInput, messageText, messageTitle, strRunChrome, oExec

messageText = "Before open chrome without security we must close all chrome.exe processes. Are you OK with this?"
messageTitle = "Chrome runner"
strRunChrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --args --disable-web-security http://localhost:8080/index.html"

Set regexNotRunning = New RegExp
regexNotRunning.Pattern = "^INFO:\s+No\s+tasks\s+are\s+running"
Set oShell = WScript.CreateObject("WScript.Shell")
Set oRet = oShell.Exec("TASKLIST /FI ""IMAGENAME eq chrome.exe""")
strCommandResult = oRet.StdOut.ReadAll()

If regexNotRunning.Test( strCommandResult )  Then
	oShell.Exec(strRunChrome)
Else
    userInput = MsgBox (messageText, vbYesNo, messageTitle)
	If userInput = vbYes Then
		Set oExec = oShell.Exec("TASKKILL /IM chrome.exe")
		Do While oExec.Status = 0
			WScript.Sleep 100
		Loop
		oShell.Exec(strRunChrome)
	End If
End If

Set oShell = Nothing
Set oRet = Nothing
Set regexNotRunning = Nothing
Set strCommandResult = Nothing
Set userInput = Nothing
Set messageText = Nothing
Set messageTitle = Nothing 
Set strRunChrome  = Nothing 
Set oExec = Nothing
