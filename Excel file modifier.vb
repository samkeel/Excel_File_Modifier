Sub defaultBehaviourOn()
  ' turn on excel defaults
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


Sub defaultBehaviourOff()
    ' disable excel defaults to speed up processing
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
End Sub

Function selectFilePath(path1 As String, dialogText As String) As String
Dim refSht As Variant
Dim fileNamePrompt

' dialog box to choose the document to import
Set refSht = Application.FileDialog(3)
    ' set deault file location for open dialog to appear in
    refSht.Title = dialogText
    refSht.InitialFileName = Application.ThisWorkbook.Path + "\" + fileNamePrompt
    refSht.AllowMultiSelect = False
    If refSht.Show = -1 Then
        path1 = refSht.SelectedItems.Item(1)
    Else
        path1 = ""
    End If

End Function