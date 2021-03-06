Attribute VB_Name = "SaveAttachment"
Option Explicit

'***********************************************************************
'* Code based on sample code from Martin Green and adapted to my needs
'* more on TheTechieGuy.com - Liron@TheTechieGuy.com
'* adapted further by Christoffer Brob�ck
'***********************************************************************

Public Sub GetAttachments()
    On Error Resume Next
    'create the folder if it doesnt exists:
    Dim fso As Variant
    Dim ttxtfile As Variant
    Dim txtfile As Variant
    Dim WheretosaveFolder As Variant

    Dim objFolders As Object
    Set objFolders = CreateObject("WScript.Shell").SpecialFolders
 
    'MsgBox objFolders("mydocuments")
    ttxtfile = objFolders("desktop")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Changes made by Andrew Davis (adavis@xtheta.com) on October 28th 2015
    ' ------------------------------------------------------
    ' Set fso = Nothing
    ' ------------------------------------------------------

    
    On Error GoTo GetAttachments_err
    ' Declare variables
    Dim ns As NameSpace
    Dim Inbox As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    Dim filename As String
    Dim i As Long
    Set ns = GetNamespace("MAPI")
    'Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    ' added the option to select whic folder to export
    Set Inbox = ns.PickFolder
    Set txtfile = fso.CreateFolder(ttxtfile & "\" & Inbox)
    WheretosaveFolder = ttxtfile & "\" & Inbox
    
    'to handle if the use cancalled folder selection
    If Inbox Is Nothing Then
        MsgBox "You need to select a folder in order to save the attachments", vbCritical, _
               "Export - Not Found"
        Exit Sub
    End If

    ''''
    

    i = 0
    ' Check Inbox for messages and exit of none found
    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the selected folder.", vbInformation, _
               "Export - Not Found"
        Exit Sub
    End If
    ' Check each message for attachments
    For Each Item In Inbox.Items
        ' Save any attachments found
        For Each Atmt In Item.Attachments
            ' This path must exist! Change folder name as necessary.
        
            ' Changes made by Andrew Davis (adavis@xtheta.com) on October 28th 2015
            ' ------------------------------------------------------
            filename = WheretosaveFolder & "\" & fso.GetBaseName(Atmt.filename) & i & "." & fso.GetExtensionName(Atmt.filename)
            ' ------------------------------------------------------
            Atmt.SaveAsFile filename
            i = i + 1
        Next Atmt
    Next Item
    ' Show summary message
    If i > 0 Then
        MsgBox "There were " & i & " attached files." _
             & vbCrLf & "These have been saved to the Email Attachments folder in My Documents." _
             & vbCrLf & vbCrLf & "Thank you for using Liron Segev - TheTechieGuy's utility", vbInformation, "Export Complete"
    Else
        MsgBox "There were no attachments found in any mails.", vbInformation, "Export - Not Found"
    End If
    ' Changes made by Andrew Davis (adavis@xtheta.com) on October 28th 2015
    ' ------------------------------------------------------
    Set fso = Nothing
    ' ------------------------------------------------------
    ' Clear memory
GetAttachments_exit:
    Set Atmt = Nothing
    Set Item = Nothing
    Set ns = Nothing
    Exit Sub
    ' Handle errors
GetAttachments_err:
    MsgBox "An unexpected error has occurred." _
         & vbCrLf & "Please note and report the following information." _
         & vbCrLf & "Macro Name: GetAttachments" _
         & vbCrLf & "Error Number: " & Err.Number _
         & vbCrLf & "Error Description: " & Err.Description _
           , vbCritical, "Error!"
    Resume GetAttachments_exit
End Sub


