Dim old_ORM As String
Dim old_launchStatus As String
Dim old_crewStatus As String
Dim old_areaStatus As String
Dim old_remarks As String
Dim old_SPLAStatus As String

Public alreadyExported As Boolean

Sub ResetStatus_Click()
    Dim cfmResponse As Variant
    cfmResponse = MsgBox("Resetting the status will delete the update log and restore system defaults. Are you sure you want to continue?", vbExclamation + vbYesNo, "Confirm Reset")
    
    If cfmResponse <> vbYes Then Exit Sub
    
    ResetStatus
End Sub
Sub ResetStatus()
    'Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    With Worksheets("Main")
        .Range("MedItems").ClearContents
        .Range("HighItems").ClearContents
        .Range("LaunchStatus").MergeArea.Value = "Approved for Launch"
        .Range("AsOfTime").MergeArea.Value = Format(Now, "hhmm")
        .Range("CrewStatus").MergeArea.Value = "Dual"
        .Range("AreaStatus").MergeArea.Value = "Both Areas"
        .Range("Remarks").MergeArea.ClearContents
        .Range("SPLAStatus").MergeArea.Value = "In Use"
        '.Range("SystemStatus").Value = "Ready for initial SCO status input."
        .Range("SystemStatus").Font.Bold = False
        .Range("SystemStatus").Offset(0, -1).Interior.Color = RGB(221, 235, 247)
        
        .Buttons("updateBtn").Font.FontStyle = "Normal"
        .Buttons("updateBtn").Visible = False
    End With
    
    With Worksheets("Log")
        .Cells.ClearContents
        .Range("A1").Value = "SCO ORM Status Log"
        .Range("A2").Value = "=TODAY()"
        .Range("A3").Value = "Time"
        .Range("B3").Value = "Category"
        .Range("C3").Value = "Event"
    End With
    
    Application.ScreenUpdating = True
    Worksheets("Main").Range("SystemStatus").Value = "Waiting for initial status..."
    'Application.DisplayAlerts = True
    
    old_ORM = ""
    old_launchStatus = ""
    old_crewStatus = ""
    old_areaStatus = ""
    old_remarks = ""
    old_SPLAStatus = ""
    
    alreadyExported = False
End Sub

Sub UpdateStatus()
    Dim newORM As String, newLaunch As String, newCrew As String, newArea As String, newRemarks As String, newSPLA As String
    
    newORM = "SCO ORM: " & Range("SCONum").Value & "." & vbNewLine
    
    Dim cell As Variant
    
    For Each cell In Worksheets("Main").Range("highItems")
        If cell.Value <> "" Then
            newORM = newORM & " ** 2 (HIGH) for " & cell.Offset(0, -4).Value & "." & vbNewLine
        End If
    Next cell
    
    For Each cell In Worksheets("Main").Range("medItems")
        If cell.Value <> "" Then
            newORM = newORM & " - 1 for " & cell.Offset(0, -2).Value & "." & vbNewLine
        End If
    Next cell
    
    ' Debugging.
    ' MsgBox "Old ORM: " & old_ORM & vbNewLine & vbNewLine & "New ORM: " & newORM & vbNewLine & vbNewLine & "Changed? >>" & (old_ORM <> newORM)
    
    With Worksheets("Main")
        newLaunch = .Range("LaunchStatus").Value
        .Range("AsOfTime").Value = Format(Now, "hhmm")
    
        newCrew = .Range("CrewStatus").Value
        newArea = .Range("AreaStatus").Value
        newRemarks = .Range("Remarks").Value
        newSPLA = .Range("SPLAStatus").Value
    End With
    
    setStatus newORM, newLaunch, newCrew, newArea, newRemarks, newSPLA
    
    With Worksheets("Main").Range("SystemStatus")
        .Value = "Update logged. Ready."
        .Font.Bold = False
        .Offset(0, -1).Interior.Color = RGB(221, 235, 247)
    End With
    
    With Worksheets("Main").Buttons("UpdateBtn")
        .Font.FontStyle = Normal
        .Visible = False
    End With
End Sub

Function getStatus() As String()
    Dim ORM As String, launchStatus As String, crewStatus As String, areaStatus As String, remarks As String, SPLAStatus As String
    
    ORM = old_ORM
    launchStatus = old_launchStatus
    crewStatus = old_crewStatus
    areaStatus = old_areaStatus
    remarks = old_remarks
    SPLAStatus = old_SPLAStatus
    
    getStatus = Array(ORM, launchStatus, crewStatus, areaStatus, remarks, SPLAStatus)
End Function

Sub setStatus(ByVal new_ORM As String, ByVal new_launchStatus As String, ByVal new_crewStatus As String, ByVal new_areaStatus As String, ByVal new_remarks As String, ByVal new_SPLAStatus As String)
    If new_ORM <> old_ORM Then
        writeLog "ORM", new_ORM
    End If
    
    If new_launchStatus <> old_launchStatus Then
        writeLog "Launch Status", new_launchStatus
    End If
    
    If new_crewStatus <> old_crewStatus Then
        writeLog "Crew Status", new_crewStatus
    End If
    
    If new_areaStatus <> old_areaStatus Then
        writeLog "Area Status", new_areaStatus
    End If
    
    If new_remarks <> old_remarks Then
        writeLog "Remarks", new_remarks
    End If
    
    If new_SPLAStatus <> old_SPLAStatus Then
        writeLog "SPLA", new_SPLAStatus
    End If
    
    old_ORM = new_ORM
    old_launchStatus = new_launchStatus
    old_crewStatus = new_crewStatus
    old_areaStatus = new_areaStatus
    old_remarks = new_remarks
    old_SPLAStatus = new_SPLAStatus

End Sub

Private Sub writeLog(ByVal logCategory As String, ByVal logText As String)
    Dim emptyRow As Integer
    Dim timeStamp As Date
    
    emptyRow = 1
    
    While Worksheets("Log").Cells(emptyRow, 1) <> ""
        emptyRow = emptyRow + 1
    Wend
    
    timeStamp = Now
    
    With Worksheets("Log")
        .Cells(emptyRow, 1).Value = timeStamp
        .Cells(emptyRow, 2).Value = logCategory
        .Cells(emptyRow, 3).Value = logText
        .Range("C1:C99999").Columns.AutoFit
    End With
End Sub

Sub ExportStatus()
    Dim tempWB As Workbook
    
    Application.ScreenUpdating = False
    
    Set tempWB = Application.Workbooks.Add
    
    ThisWorkbook.Worksheets("Log").Copy Before:=tempWB.Sheets(tempWB.Sheets.Count)
    
    tempWB.SaveAs Filename:=ThisWorkbook.Path & "\" & "Log_" & Format(Now, "dd_mmm_yy") & ".csv", FileFormat:=xlCSV
    
    tempWB.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    alreadyExported = True
    
End Sub


