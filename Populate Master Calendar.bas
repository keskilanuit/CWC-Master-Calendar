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
    
  
    Dim Cell As Range
    Dim formattingRange As Range
 
    Dim exampleRow As Range
    Dim taargetRange As Range
    
    Dim lastUsedRow As Long
    Dim lastRowToDelete As Long
    

    Set ws = ThisWorkbook.Worksheets("Master Calendar")
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Define the example row range (rows 3 and 4) with formatting you want to apply
    Set exampleRow = ws.Range("A3:M4")
    
    ' Define the target range (rows 5 to lastRow) for formatting
    Set targetRange = ws.Range("A5:M" & lastRow)
    
    ' Apply the formatting from the example row to the target range
    exampleRow.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    
      
      
    Set matrixSheet = ThisWorkbook.Worksheets("Matrix")
    Set masterSheet = ThisWorkbook.Worksheets("Master Calendar")
    
    
    
     Worksheets("Master Calendar").Range("A3:Z" & Rows.Count).ClearContents
    
    lastRow = matrixSheet.Cells(matrixSheet.Rows.Count, "M").End(xlUp).Row
      
    
    
    For i = 3 To lastRow
        If matrixSheet.Range("M" & i).Value = "X" Then
            masterLastRow = masterSheet.Cells(masterSheet.Rows.Count, "A").End(xlUp).Row
            matrixSheet.Range("A" & i & ":L" & i).Copy
            masterSheet.Range("A" & masterLastRow + 1).PasteSpecial xlPasteValues
             masterSheet.Range("A" & masterLastRow + 1).PasteSpecial xlPasteFormats
            

            If masterSheet.Range("N" & masterLastRow + 1).Value = "" Then
                masterSheet.Range("M" & masterLastRow + 1).Value = "Not Start"
             numRowsCopied = numRowsCopied + 1
                    End If
        End If
    Next i
        
        
        
        
            ' Add a blank row after pasting the last row from the initial VBA code
    masterLastRow = masterSheet.Cells(masterSheet.Rows.Count, "A").End(xlUp).Row
    masterSheet.Rows(masterLastRow + 1).Insert Shift:=xlDown

 
    With masterSheet.Range("A" & masterLastRow + 1)
        .HorizontalAlignment = xlLeft
        .Value = "PeriodEnd task items shown below"
        .Font.Size = 24
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0) ' RED font color
        End With
        
            With masterSheet.Range("A" & masterLastRow + 1 & ":L" & masterLastRow + 1)
        .Interior.Color = RGB(255, 255, 0) ' Yellow background color
    End With
        
        

    
        
        

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
    


    Set ws = ThisWorkbook.Worksheets("Master Calendar")

    ' Find the last used row in column A starting from cell A3
    lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Calculate the row number to delete (last used row + 1 to 50 rows after that)
    lastRowToDelete = lastUsedRow + 1 + 50

    ' Loop through and delete the next 50 unused rows from the last used row
    For i = lastRowToDelete To lastUsedRow + 1 Step -1
        ws.Rows(i).Delete
    Next i


      MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 MsgBox "Finished populating Current month's Master Calendar, total of " & numRowsCopied & " returns need to be prepare for this period."
  MsgBox "This code ran successfully in " & MinutesElapsed & " seconds", vbInformation
 

 
End Sub
