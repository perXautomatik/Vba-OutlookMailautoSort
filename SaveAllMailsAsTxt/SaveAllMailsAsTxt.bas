Attribute VB_Name = "SaveAllMailsAsTxt"
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
    Dim DiarieSet
    Dim FastighetSet
    Dim RetStr As String
    
    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.PickFolder
    WheretosaveFolder = ttxtfile & "\" & Inbox
    
    
    If Inbox Is Nothing Then
                MsgBox "You need to select a folder in order to save the attachments", vbCritical, _
               "Export - Not Found"
        GoTo LastLine
    End If

    i = 0

    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the selected folder.", vbInformation, _
               "Export - Not Found"
        GoTo LastLine
    End If

    FileName = WheretosaveFolder & ".txt"
    Set objFile = objFSO.CreateTextFile(FileName, True)

    For Each objItem In Inbox.Items
           
        With CreateObject("vbscript.regexp")
            .Pattern = "\<.*?\>"
            .Global = True
            NoLineBreaksNoHtml = .Replace(Replace(Replace(Replace(Replace(objItem.HTMLBody & "~" & objItem.Subject, Chr(10), ""), vbCrLf, " "), vbLf, " "), vbCr, " "), "")
        End With
        
            Set objRegExp = New RegExp
            objRegExp.IgnoreCase = True
            objRegExp.Global = True
        DiarieSet = Array()
        If (RxMatch(NoLineBreaksNoHtml, "[MHNmhnbBVv]{1,4}[-]\d{4}[-]\d{1,4}\s")) Then
                Call unique(RxMatches(NoLineBreaksNoHtml, "[MHNmhnbBVv]{1,4}[-]\d{4}[-]\d{1,4}\s"), Udiarie)
        End If
        
        FastighetSet = Array()
        
        If (RxMatch(NoLineBreaksNoHtml, "[^\s\d]{0,}\s?[^\s\d]{1,}\s[sS\d]{1,4}[:]\d{1,4}\s")) Then
                 Call unique(RxMatches(NoLineBreaksNoHtml, "[^\s\d]{0,}\s?[^\s\d]{1,}\s[sS\d]{1,4}[:]\d{1,4}\s"), UFastighet)
        End If
                
    Dim var1 As Variant
    Dim var2 As Variant
    Dim var3 As Variant
    
    var1 = objItem.entryId
    
    If (IsArrayEmpty(Udiarie) >= 1) Then
        var2 = Udiarie(1)
    End If
    
    If (IsArrayEmpty(UFastighet) >= 1) Then
        var3 = UFastighet(1)
    End If
    
    Debug.Print (var2)
        Debug.Print ("var2")
    Debug.Print (var3)
        Debug.Print ("var3")
    objFile.writeline (var1 & "~" & var2 & "~" & var3)

        Set Udiarie = Nothing
        Set UFastighet = Nothing
        Set DiarieSet = Nothing
        Set UFastighet = Nothing
    Next objItem
        Set fso = Nothing
        

LastLine:
' Clear memory
SaveMailAs_exit:
    Set Atmt = Nothing
    Set objItem = Nothing
    Set ns = Nothing
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

Public Function IsArrayEmpty(Arr As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayEmpty
    ' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
    '
    ' The VBA IsArray function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This function tests whether the array has actually
    ' been allocated.
    '
    ' This function is really the reverse of IsArrayAllocated.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If
    
    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occassion, under circumstances I
        ' cannot reliably replictate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBoung is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
    
    
    End Function

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
    Dim arrMatches
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

Sub unique(duped As Variant, unduped As Variant)
    
        
Debug.Print "before" & UBound(duped) - LBound(duped) + 1
    
    Dim a As Variant

    On Error Resume Next
    For Each a In duped
         unduped.Add a.Value, a.Value
  Next

Debug.Print "after" & UBound(unduped) - LBound(unduped) + 1

End Sub






