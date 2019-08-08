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
    Dim DiarieSet   As MatchCollection
    Dim FastighetSet   As MatchCollection
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
            Debug.Print (NoLineBreaksNoHtml)
        End With
        
        Set objRegExp = New RegExp
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
    
        objRegExp.Pattern = "^[^-]*([MHNmhnbBVvMBNV]{1,4}[-]\d{4}[-]\d{1,4}\s)"
       
        If (objRegExp.Test(NoLineBreaksNoHtml) = True) Then
            Set DiarieSet = objRegExp.Execute(NoLineBreaksNoHtml)
        End If
    
        objRegExp.Pattern = "[, ]{2}([^\d]*\d{1,2}[:]\d{1,2})\s?$"
        If (objRegExp.Test(NoLineBreaksNoHtml) = True) Then
             Set FastighetSet = objRegExp.Execute(NoLineBreaksNoHtml)
        End If
        
        'Call unique(DiarieSet, Udiarie)
        'Call unique(FastighetSet, UFastighet)
                
    Dim var1 As Variant
    Dim var2 As Variant
    Dim var3 As Variant
    
    var1 = objItem.entryId
    Debug.Print (Udiarie.Count & var1)
    Debug.Print (UFastighet.Count)
    
    If (IsArrayEmpty(Udiarie) >= 1) Then
        var2 = Udiarie(1)
        Debug.Print (var2)
        Debug.Print ("var2")
        
    End If
    
    If (IsArrayEmpty(UFastighet) >= 1) Then
        var3 = UFastighet(1)
        Debug.Print (var3)
        Debug.Print ("var3")
    End If

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

Function Regex(MyString As String, MyPattern As String) As String
    Dim Regex As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = "^[0-9]{1,3}"

    If strPattern <> "" Then
        strInput = Myrange.Value
        strReplace = ""

        With Regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If Regex.Test(strInput) Then
            simpleCellRegex = Regex.Replace(strInput, strReplace)
        Else
            simpleCellRegex = "Not matched"
        End If
    End If
End Function

Sub unique(duped As MatchCollection, unduped As Collection)

Dim a As Variant

  On Error Resume Next
  For Each a In duped
     unduped.Add a.Value, a.Value
  Next

End Sub






