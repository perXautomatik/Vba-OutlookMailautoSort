Attribute VB_Name = "organizeMailsBySocken"
Option Explicit '   // Force: Declare your Variables
Public Sub Move_Items()

    Dim Inbox As Outlook.MAPIFolder

    Dim olNs As Outlook.NameSpace
    Dim Item As Object
    Dim dataline As String
    
    Set olNs = Application.GetNamespace("MAPI")
    Set Inbox = olNs.PickFolder
   
    Open "C:\Users\crbk01\Desktop\ToExport.csv" For Input As #2
    
    On Error GoTo ErrorHandler
        While Not EOF(2)
            Line Input #2, dataline
            
            Dim LineItems() As String
            Dim värde As Long
            Dim entryId As String
            LineItems = Split(dataline, ",")
            Dim SubFolder As Outlook.MAPIFolder
            
            If (UBound(LineItems) = 1) Then
                värde = Replace(LineItems(1), Chr$(34), vbNullString)
                entryId = Replace(LineItems(0), Chr$(34), vbNullString)
                
                Set Item = olNs.GetItemFromID(entryId, Inbox.StoreID)
                Item.UnRead = False
                
                If TypeName(Item) <> "Nothing" Then
                    värde = IIf(värde > 3, värde - 1, värde)
                    värde = IIf(värde = 0, 6, värde)
                   
                    With Inbox
                    Select Case värde
                    
                        Case 1
                            Set SubFolder = .Folders("södra")
                        Case 2
                            Set SubFolder = .Folders("Norra")
                        Case 3
                            Set SubFolder = .Folders("mellersta")
                        Case 4
                            Set SubFolder = .Folders("distrikt")
                        Case 5
                            Set SubFolder = .Folders("kanske gk ejuts")
                        Case 6
                            Set SubFolder = .Folders("Ansökningar, ej sorterade")
                        End Select
                        
                        Item.Move SubFolder
                            
                    End With
                    
                End If
            End If
        Wend
        

    Set Inbox = Nothing
    Set SubFolder = Nothing
    Set olNs = Nothing
    Set Item = Nothing
    Close #2

    Exit Sub

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub
