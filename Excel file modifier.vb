Sub fileModifier()
Dim refOpenWB1 As Workbook
Dim refOpenWS1 As Worksheet
Dim refOpenWB2 As Workbook
Dim refOpenWS2 As Worksheet
Dim dialogText As String
Dim path1 As String
Dim path2 As String
Dim datarange1 As Variant
Dim datarange2 As Variant
Dim i As Long
Dim j As Long
Dim curRow As Long

dialogText = "Document1"
selectFilePath path1, dialogText
If path1 = "" Then
    GoTo fileOpenCancell
End If
'Document 1
'Read contents of sheet into single variable - datarange1
Set refOpenWB1 = Workbooks.Open(path1, UpdateLinks:=False, ReadOnly:=True)
WorkBook1Name = refOpenWB1.Name
Set refOpenWS1 = refOpenWB1.Sheets("Sheet1")
datarange1 = refOpenWS1.Range("A6").CurrentRegion.Value
refOpenWB1.Close (False)

dialogText = "Document2"
selectFilePath path2, dialogText
If path2 = "" Then
    GoTo fileOpenCancell
End If
'Document 2
'Read contents of sheet into single variable - datarange2
Set refOpenWB2 = Workbooks.Open(path2, UpdateLinks:=False, ReadOnly:=True)
WorkBook2Name = refOpenWB2.Name
Set refOpenWS2 = refOpenWB2.Sheets("Sheet1")
datarange2 = refOpenWS2.Range("A6").CurrentRegion.Value
refOpenWB2.Close (False)

Call defaultBehaviourOff

'-------
'Loop through dataset1 and dataset2 making any changes
'-------

For i = 1 To UBound(datarange1) - LBound(datarange1) + 1
    For j = 1 To UBound(datarange2) - LBound(datarange2) + 1
'--- Add code here
    Next j
Next i

'Open Document 1 again, this time in write mode to save any changes.
 Set refOpenWB1 = Workbooks.Open(path1, UpdateLinks:=False, ReadOnly:=False)
 WorkBook1Name = refOpenWB1.Name
 Set refOpenWS1 = refOpenWB1.Sheets("Sheet1")

curRow = 1

'loop through rows updating fields eg:
For i = 1 To UBound(datarange1) - LBound(datarange1) + 1
    'change a cells text to red to show it is updated.
    With refOpenWS1.Cells(curRow, 2)
        .Font.Color = vbRed
    End With
    curRow = curRow + 1
Next i

'Close and save workbook
refOpenWB1.Close (True)

'------
'- Program End and error catch
'------

'program End
Done:
  Call defaultBehaviourOn
  Exit Sub

'error catch
'File selection cancel
fileOpenCancell:
    Call defaultBehaviourOn
    MsgBox "cancelled"
    Exit Sub

End Sub

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
