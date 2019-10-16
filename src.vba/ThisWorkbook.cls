Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim cfmLog As Variant
    
    If Not alreadyExported Then
        cfmLog = MsgBox("Caution! Closing will delete the update log (unless already exported to a file). Are you sure you want to continue?", vbExclamation + vbYesNo, "Confirm Close")
    
        If cfmLog <> vbYes Then Cancel = True
    End If
End Sub

Private Sub Workbook_Open()
    ResetStatus
End Sub
