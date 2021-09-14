Combine-Macros.vbs macro
'(Copy URL in PhpBB Forum Format - Info)

'This example is for combining calls to macros in a single script. It calls the macros: Wsh-Start.iim, Wsh-Lunch.iim, and Wsh-Submit-Button. The example also shows how to create a custom error report log file. The script is also available as PowerShell version.

'Similar script: self-test.vbs

'Visual Basic Script:

' iMacros Combine-Macros Script
' (c) 2008-2010 iOpus Software 

'*****
'This part is for the script specific error log (instead of popup messages).
'The advantage of an error log instead of a popup is that the script 
'continues running even after an error appears. An example of an error would
'be a timeout due to a temporarily slow web site. 

Option Explicit

Dim objFileSystem, objOutputFile
Dim strOutputFile

Const OPEN_FILE_FOR_APPENDING = 8
' generate a logfile name based on the script name
strOutputFile = "./combine-macros-errorlog.txt" 
Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objOutputFile = objFileSystem.CreateTextFile(strOutputFile, TRUE)
objOutputFile.WriteLine("Error Log for COMBINE-MACROS.VBS demo script")

'****
' find current folder
Dim myname, mypath
myname = WScript.ScriptFullName
mypath = Left(myname, InstrRev(myname, "\"))

Dim message
message = "This example script calls several Internet macros. Each macro performs a specific function on the website "
message = message + "(Loading the website with wsh-start, form filling with wsh-lunch and finally submitting the form with wsh-submit)." + vbNewLine
message = message + "The script also demonstrates how to create a SCRIPT SPECIFIC error log file."

Dim iim1, i
set iim1 = CreateObject ("imacros")

i =  iim1.iimOpen
'If i < 0 Then msgbox iim1.iimGetErrorText()
if i < 0 then objOutputFile.WriteLine("INIT: Error-No: " + cstr(i) + " => Description: " + iim1.iimGetErrorText())


i = iim1.iimPlay(mypath & "Macros\wsh-start.iim")
'If i < 0 Then msgbox iim1.iimGetErrorText()
if i < 0 then objOutputFile.WriteLine("WSH-START: Error-No: " + cstr(i) + " => Description: " + iim1.iimGetErrorText())


i = iim1.iimPlay(mypath & "Macros\wsh-lunch.iim")
'If i < 0 Then msgbox iim1.iimGetErrorText()
if i < 0 then objOutputFile.WriteLine("WSH-LUNCH: Error-No: " + cstr(i) + " => Description: " + iim1.iimGetErrorText())


i = iim1.iimPlay(mypath & "Macros\wsh-submit-button.iim")
'If i < 0 Then msgbox iim1.iimGetErrorText()
if i < 0 then objOutputFile.WriteLine("WSH-SUBMIT: Error-No: " + cstr(i) + " => Description: " + iim1.iimGetErrorText())


i = iim1.iimClose
if i < 0 then objOutputFile.WriteLine("EXIT: Error-No: " + cstr(i) + " => Description: " + iim1.iimGetErrorText())


'*****

'This part is for the script specific error log (instead of popup messages).
objOutputFile.Close
Set objFileSystem = Nothing

'*****

WScript.Quit(i)

Wsh-Start.iim macro

TAB T=1     
TAB CLOSEALLOTHERS  
URL GOTO=http://demo.imacros.net/Automate/     
TAG POS=1 TYPE=A ATTR=TXT:*Testform1*   
TAG POS=1 TYPE=INPUT:TEXT FORM=ID:demo ATTR=NAME:name CONTENT=Tom<SP>Tester 
TAG POS=1 TYPE=SELECT FORM=ID:demo ATTR=NAME:food CONTENT=$Pizza
TAG POS=1 TYPE=SELECT FORM=ID:demo ATTR=NAME:drink CONTENT=$Coke
TAG POS=1 TYPE=INPUT:RADIO FORM=ID:demo ATTR=ID:small&&VALUE:small CONTENT=YES 
TAG POS=1 TYPE=SELECT FORM=ID:demo ATTR=ID:dessert CONTENT=%ice<SP>cream:%apple<SP>pie
TAG POS=1 TYPE=INPUT:RADIO FORM=ID:demo ATTR=NAME:Customer CONTENT=YES
TAG POS=1 TYPE=INPUT:RADIO FORM=ID:demo ATTR=NAME:Customer&&VALUE:Not_yet CONTENT=YES 

Wsh-Lunch.iim macro

TAG POS=1 TYPE=TEXTAREA FORM=ID:demo ATTR=NAME:Remarks CONTENT=Lunch

Wsh-Submit-Button.iim macro

TAB T=1     
TAB CLOSEALLOTHERS   
TAG POS=1 TYPE=BUTTON:SUBMIT FORM=ID:demo ATTR=TXT:Click<SP>to<SP>order<SP>now
WAIT SECONDS=3
URL GOTO=http://demo.imacros.net/Automate/OK

