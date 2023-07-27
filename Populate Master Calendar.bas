Attribute VB_Name = "Module2"
Sub Populate_Master_Calendar()

'Author: Keski Lin'

    answer = MsgBox("Executing this macro may take up to 120 minutes, failed on running macro may result in crash all opened excel files, please make sure all your workbooks are saved. Do you wish to continue?", vbQuestion + vbYesNo, "Warning")
    
      If answer = vbYes Then

    Else
        Exit Sub
    End If
    

    Dim StartTime As Double
    Dim MinutesElapsed As String

     StartTime = Timer
     
    
    Dim matrixSheet As Worksheet, masterSheet As Worksheet
    Dim lastRow As Long, i As Long, masterLastRow As Long
    
    Set matrixSheet = ThisWorkbook.Worksheets("Matrix")
    Set masterSheet = ThisWorkbook.Worksheets("Master Calendar")
    
     Worksheets("Master Calendar").Range("A3:Z" & Rows.Count).ClearContents
    
    lastRow = matrixSheet.Cells(matrixSheet.Rows.Count, "M").End(xlUp).Row
    
    For i = 3 To lastRow
        If matrixSheet.Range("M" & i).Value = "X" Then
            masterLastRow = masterSheet.Cells(masterSheet.Rows.Count, "A").End(xlUp).Row
            matrixSheet.Range("A" & i & ":L" & i).Copy
            masterSheet.Range("A" & masterLastRow + 1).PasteSpecial xlPasteValues
            

            If masterSheet.Range("N" & masterLastRow + 1).Value = "" Then
                masterSheet.Range("M" & masterLastRow + 1).Value = "Not Start"
             numRowsCopied = numRowsCopied + 1
                    End If
        End If
    Next i
        

      MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 MsgBox "Finished populating Current month's Master Calendar, total of " & numRowsCopied & " returns need to be prepare for this period."
  MsgBox "This code ran successfully in " & MinutesElapsed & " seconds", vbInformation
 

 
End Sub
