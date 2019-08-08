Attribute VB_Name = "SaveAllMailsAsTxt1"
Option Explicit

'***********************************************************************
'* Code based on sample code from Martin Green and adapted to my needs
'* more on TheTechieGuy.com - Liron@TheTechieGuy.com
'* adapted further by Christoffer Brobäck
'***********************************************************************

Sub SaveMailAs()
    Dim fso, ttxtfile, txtfile, WheretosaveFolder
    Dim objFolders, objFSO As Object
    Set objFolders = CreateObject("WScript.Shell").SpecialFolders
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'MsgBox objFolders("mydocuments")
    ttxtfile = objFolders("desktop")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ns As NameSpace
    Dim Inbox As MAPIFolder
    Dim objItem As Object
    Dim objFile As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim i As Integer
    Dim regId As String
    Dim xcell As Variant
    Dim NoLineBreaksNoHtml As Variant
    Dim diarie As Collection
    Dim fastighet As Collection
    Dim Udiarie As New Collection
    Dim UFastighet As New Collection
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim DiarieSet   As MatchCollection
    Dim FastighetSet   As MatchCollection
    Dim RetStr As String
    
    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.PickFolder
    WheretosaveFolder = ttxtfile & "\" & Inbox
    
    If Inbox Is Nothing Then
                MsgBox "You need to select a folder in order to save the attachments", vbCritical, _
               "Export - Not Found"
        Exit Sub
    End If

    i = 0

    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the selected folder.", vbInformation, _
               "Export - Not Found"
        Exit Sub
    End If

    FileName = WheretosaveFolder & ".txt"
    Set objFile = objFSO.CreateTextFile(FileName, False)

    For Each objItem In Inbox.Items
           
        With CreateObject("vbscript.regexp")
            .Pattern = "\<.*?\>"
            .Global = True
               NoLineBreaksNoHtml = .Replace(Replace(Replace(Replace(Replace(objItem.HTMLBody & "~" & objItem.Subject, Chr(10), ""), vbCrLf, " "), vbLf, " "), vbCr, " "), "")
        End With
        
            Set objRegExp = New RegExp
            objRegExp.IgnoreCase = True
            objRegExp.Global = True
        
           objRegExp.Pattern = "[MHNmhnbBVv]{1,4}[-]\d{4}[-]\d{1,4}\s"
           
           If (objRegExp.Test(NoLineBreaksNoHtml) = True) Then
                     Set DiarieSet = objRegExp.Execute(NoLineBreaksNoHtml)
            End If
        
            objRegExp.Pattern = "[^\s\d]{0,}\s?[^\s\d]{1,}\s[sS\d]{1,4}[:]\d{1,4}\s"
           If (objRegExp.Test(NoLineBreaksNoHtml) = True) Then
                 Set FastighetSet = objRegExp.Execute(NoLineBreaksNoHtml)
            End If
        
        Call unique(DiarieSet, Udiarie)
        Call unique(FastighetSet, UFastighet)
                
        objFile.writeline (objItem.entryId & "~" & IIf(Udiarie.Count >= 1, Udiarie(1), "") & "~" & IIf(UFastighet.Count >= 1, UFastighet(1), ""))

'path = windows temp & todaysDate()
    
       Set Udiarie = Nothing
        Set UFastighet = Nothing
        Set DiarieSet = Nothing
        Set UFastighet = Nothing
    Next objItem
        Set fso = Nothing
        

' Clear memory
SaveMailAs_exit:
    Set Atmt = Nothing
    Set objItem = Nothing
    Set ns = Nothing
    
    Exit Sub
' Handle errors
'GetAttachments_err:
 '   MsgBox "An unexpected error has occurred." _
  '      & vbCrLf & "Please note and report the following information." _
   '     & vbCrLf & "Macro Name: GetAttachments" _
    '    & vbCrLf & "Error Number: " & Err.Number _
     '   & vbCrLf & "Error Description: " & Err.Description _
      '  , vbCritical, "Error!"
 '   Resume GetAttachments_exit
End Sub

Sub unique(duped As MatchCollection, unduped As Collection)

Dim a As Variant

  On Error Resume Next
  For Each a In duped
     unduped.Add a.Value, a.Value
  Next

End Sub






