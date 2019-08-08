Attribute VB_Name = "MainModule"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub main()

    Dim filename As String: filename = SaveAllMailsAsTxt.SaveMailAs
    
    Do
        If fso.FileExists(filename) Then
            Exit Do
        End If

        DoEvents                                 'Prevents from being unresponsive
        Sleep 1000                               '1 Second
    Loop

    Debug.Print "file available"

    
    Call organizeMailsBySocken.Move_Items
    
    Call SaveAttachment.GetAttachments
    
    Call Shell("POWERSHELL.exe -noexit " & _
              """H:\Operations\REPORTS\Reports2018\Balance Sheet\SLmarginJE.ps1""", 1)
End Sub


