Attribute VB_Name = "Modul1"
Sub SetSelectedCallendarObjectsAsFrånvarande()
    Dim Item As Object
    Dim apt As AppointmentItem
    
    For Each Item In Application.ActiveExplorer.Selection
            If Item.Class = olAppointment Then
                Set apt = Item
                apt.BusyStatus = 3
                apt.Save
            End If
    Next
End Sub
