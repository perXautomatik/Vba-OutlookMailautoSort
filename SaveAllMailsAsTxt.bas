Attribute VB_Name = "SaveAllMailsAsTxt"
Option Explicit

'***********************************************************************
'* Code based on sample code from Martin Green and adapted to my needs
'* more on TheTechieGuy.com - Liron@TheTechieGuy.com
'* adapted further by Christoffer Brobäck
'***********************************************************************

Public Sub SaveMailAs()
   
   
    Dim objFolders As Variant: Set objFolders = CreateObject("WScript.Shell").SpecialFolders
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim fso As Variant: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ns As NameSpace: Set ns = GetNamespace("MAPI")
    
    Dim ttxtfile As Variant: ttxtfile = objFolders("desktop")
    Dim Inbox As MAPIFolder: Set Inbox = ns.PickFolder
    Dim WheretosaveFolder As Variant: WheretosaveFolder = ttxtfile & "\" & Inbox
   
        
    If Inbox Is Nothing Then
        MsgBox "You need to select a folder in order to save the attachments", vbCritical, _
               "Export - Not Found"
        GoTo LastLine
    End If
    
    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the selected folder.", vbInformation, _
               "Export - Not Found"
        GoTo LastLine
    End If

    Dim filename As String: filename = WheretosaveFolder & ".txt"
    Dim objFile As Object: Set objFile = objFSO.CreateTextFile(filename, True)
    
    Dim objItem As Object: For Each objItem In Inbox.Items
        Dim NoLineBreaksNoHtml As Variant
        With CreateObject("vbscript.regexp")
            .Pattern = "\<.*?\>"
            .Global = True
            NoLineBreaksNoHtml = .Replace(Replace(Replace(Replace(Replace(objItem.HTMLBody & "~" & objItem.Subject, Chr$(10), vbNullString), vbCrLf, " "), vbLf, " "), vbCr, " "), vbNullString)
        End With
         
        Dim DiarieSet As Variant: DiarieSet = Array()
        Dim var2 As Variant
        If (RxMatch(NoLineBreaksNoHtml, "[MHNmhnbBVv]{1,4}[-]\d{4}[-]\d{1,4}\s")) Then
            DiarieSet = RxMatches(NoLineBreaksNoHtml, "[MHNmhnbBVv]{1,4}[-]\d{4}[-]\d{1,4}\s")
            var2 = DiarieSet(LBound(DiarieSet))
        End If
        
        Dim FastighetSet As Variant: FastighetSet = Array()
        Dim var3 As Variant
        If (RxMatch(NoLineBreaksNoHtml, "[^\s\d]{0,}\s?[^\s\d]{1,}\s[sS\d]{1,4}[:]\d{1,4}\s")) Then
            FastighetSet = RxMatches(NoLineBreaksNoHtml, "[^\s\d]{0,}\s?[^\s\d]{1,}\s[sS\d]{1,4}[:]\d{1,4}\s")
            var3 = FastighetSet(LBound(FastighetSet))
        End If
              
Debug.Print (var2 & " var2")
Debug.Print (var3 & " var3")
        
        objFile.writeline (objItem.entryId & "~" & var2 & "~" & var3)
    
    Next objItem
    
        

LastLine:
    ' Clear memory
    Set fso = Nothing
    Set objItem = Nothing
    Set ns = Nothing
End Sub

Public Function RxMatch( _
       ByVal SourceString As String, _
       ByVal Pattern As String, _
       Optional ByVal IgnoreCase As Boolean = True, _
       Optional ByVal MultiLine As Boolean = True) As Boolean
 
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = False
        .Pattern = Pattern
        RxMatch = .Test(SourceString)
    End With
    
End Function

Public Function RxMatches( _
       ByVal SourceString As String, _
       ByVal Pattern As String, _
       Optional ByVal IgnoreCase As Boolean = True, _
       Optional ByVal MultiLine As Boolean = True, _
       Optional ByVal MatchGlobal As Boolean = True) As Variant
 
    Dim oMatch As Match
    Dim arrMatches As Variant
    Dim lngCount As Long
    
    ' Initialize to an empty array
    arrMatches = Array()
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = MatchGlobal
        .Pattern = Pattern
        For Each oMatch In .Execute(SourceString)
            ReDim Preserve arrMatches(lngCount)
            arrMatches(lngCount) = oMatch.Value
            lngCount = lngCount + 1
        Next
    End With
 
    RxMatches = arrMatches
End Function


