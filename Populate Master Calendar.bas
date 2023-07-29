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





' Continue to loop through "PeriodEnd Tasks" worksheet
    Set matrixSheet = ThisWorkbook.Worksheets("PeriodEnd Tasks")

    ' Find the last row in column M of the matrixSheet
    lastRow = matrixSheet.Cells(matrixSheet.Rows.Count, "M").End(xlUp).Row

    ' Set the masterSheet variable (change "Sheet2" to your master sheet name)
    Set masterSheet = ThisWorkbook.Worksheets("Master Calendar")

    ' Loop through the matrixSheet from row 3 to lastRow
    For i = 3 To lastRow
        If matrixSheet.Range("M" & i).Value = "X" Then
            ' Find the last row in column A of the masterSheet
            masterLastRow = masterSheet.Cells(masterSheet.Rows.Count, "A").End(xlUp).Row

            ' Copy values and paste to the masterSheet with formatting
            matrixSheet.Range("A" & i & ":L" & i).Copy
            masterSheet.Range("A" & masterLastRow + 1).PasteSpecial xlPasteValues
            masterSheet.Range("A" & masterLastRow + 1).PasteSpecial xlPasteFormats

            ' Convert formulas to hardcoded values in the target range
            masterSheet.Range("A" & masterLastRow + 1).Value = masterSheet.Range("A" & masterLastRow + 1).Value
            
            ' Check and update column N in the masterSheet
            If masterSheet.Range("N" & masterLastRow + 1).Value = "" Then
                masterSheet.Range("M" & masterLastRow + 1).Value = "Not Start"
                numRowsCopied = numRowsCopied + 1
            End If
        End If
    Next i
    
    ' Clear clipboard
    Application.CutCopyMode = False



        

      MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 MsgBox "Finished populating Current month's Master Calendar, total of " & numRowsCopied & " returns need to be prepare for this period."
  MsgBox "This code ran successfully in " & MinutesElapsed & " seconds", vbInformation
 

 
End Sub
