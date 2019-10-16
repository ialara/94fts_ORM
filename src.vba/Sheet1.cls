Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo Handler
    If Not (Intersect(Target, Range("medItems")) Is Nothing And Intersect(Target, Range("highItems")) Is Nothing And _
            Intersect(Target, Range("LaunchStatus").MergeArea) Is Nothing And Intersect(Target, Range("CrewStatus").MergeArea) Is Nothing And _
            Intersect(Target, Range("AreaStatus").MergeArea) Is Nothing And Intersect(Target, Range("Remarks").MergeArea) Is Nothing And _
            Intersect(Target, Range("SPLAStatus").MergeArea) Is Nothing) Then
            
        Range("SystemStatus").Value = "SCO ORM/Status changed...submit to log when set!"
        Range("SystemStatus").Font.Bold = True
        Range("SystemStatus").Offset(0, -1).Interior.ColorIndex = 6
        
        Buttons("updateBtn").Font.FontStyle = "Bold"
        Buttons("updateBtn").Visible = True
    End If
Handler:
    Exit Sub
End Sub

'Private Function AnyIntersect(ByVal rg1 As Range, ByVal rg2 As Range, Optional ByVal rg3 As Range = Range("Z999"), Optional ByVal rg4 As Range = Range("Z999"), Optional ByVal rg5 As Range = Range("Z999"), Optional ByVal rg6 As Range = Range("Z999"), Optional ByVal rg7 As Range = Range("Z999")) As Boolean
'    Dim workingBool As Boolean
'    workingBool = False
'
'    If Not Intersect(rg1, rg2) Is Nothing Then workingBool = True
'    If Not Intersect(rg1, rg3) Is Nothing Then workingBool = True
'    If Not Intersect(rg1, rg4) Is Nothing Then workingBool = True
'    If Not Intersect(rg1, rg5) Is Nothing Then workingBool = True
'    If Not Intersect(rg1, rg6) Is Nothing Then workingBool = True
'    If Not Intersect(rg1, rg7) Is Nothing Then workingBool = True
'
'    If Intersect(rg2, rg3) Is Not Nothing Then workingBool = True
'    If Intersect(rg2, rg4) Is Not Nothing Then workingBool = True
'    If Intersect(rg2, rg5) Is Not Nothing Then workingBool = True
'    If Intersect(rg2, rg6) Is Not Nothing Then workingBool = True
'    If Intersect(rg2, rg7) Is Not Nothing Then workingBool = True
'
'    If Intersect(rg3, rg4) Is Not Nothing Then workingBool = True
'    If Intersect(rg3, rg5) Is Not Nothing Then workingBool = True
'    If Intersect(rg3, rg6) Is Not Nothing Then workingBool = True
'    If Intersect(rg3, rg7) Is Not Nothing Then workingBool = True
'
'    If Intersect(rg4, rg5) Is Not Nothing Then workingBool = True
'    If Intersect(rg4, rg6) Is Not Nothing Then workingBool = True
'    If Intersect(rg4, rg7) Is Not Nothing Then workingBool = True
'
'    If Intersect(rg5, rg6) Is Not Nothing Then workingBool = True
'    If Intersect(rg5, rg7) Is Not Nothing Then workingBool = True
'
'    If Intersect(rg6, rg7) Is Not Nothing Then workingBool = True
'
'    AnyIntersect = workingBool
'End Function