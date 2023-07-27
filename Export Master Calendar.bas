Attribute VB_Name = "Module1"
Sub Export_Master_Calendar()


'Author: Keski Lin'

    Dim masterSheet As Worksheet
    Dim lastRow As Long, i As Long, masterLastRow As Long
    Dim exportName As String, exportFolder As String
    Dim exportPath As String, exportFile As String
    Dim overwriteAnswer As VbMsgBoxResult
    
    Set masterSheet = ThisWorkbook.Worksheets("Master Calendar")
    lastRow = masterSheet.Cells(masterSheet.Rows.Count, "A").End(xlUp).Row
    

    Dim priorMonthYear As String, priorMonthMonth As String
    priorMonthYear = Year(Date)
    priorMonthMonth = Format(Date, "MM")
    

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select export folder"
        If .Show = -1 Then
            exportFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    exportName = "Master Calendar - " & priorMonthYear & "-" & priorMonthMonth & ".xlsx"
    exportPath = exportFolder & "\" & exportName
    

    If Dir(exportPath) <> "" Then
        overwriteAnswer = MsgBox("The file " & exportName & " already exists in the selected folder. Do you want to overwrite it?", vbQuestion + vbYesNo, "File already exists")
        If overwriteAnswer = vbNo Then Exit Sub
    End If
    

    Application.DisplayAlerts = False
    masterSheet.Copy
    ActiveWorkbook.SaveAs fileName:=exportPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    MsgBox "Master Calendar has been exported successfully to your selected folder."


End Sub
