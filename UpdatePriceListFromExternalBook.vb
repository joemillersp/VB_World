Sub UpdatePriceListFromExternalBook()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim referenceBook As Workbook
    Dim lastRow1 As Long, lastRow2 As Long
    Dim cell1 As Range, r2 As Range, headerCell As Range
    Dim itemNumber As String
    Dim logFile As Integer
    Dim filePath As String
    Dim found As Boolean
    Dim filePicker As Variant
    Dim ciDirectCol As Long
    Dim headerFoundCount As Integer
    Dim searchRange As Range
    
    On Error GoTo ErrorHandler
    
    ' Set the path for the log file
    filePath = "LogFile.txt"
    logFile = FreeFile()
    
    ' Open the log file for writing
    Open filePath For Append As #logFile
    
    ' Open the file dialog box for the user to select a file
    filePicker = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Select File")

    ' Check if the user canceled the dialog
    If filePicker <> False Then
        Set referenceBook = Workbooks.Open(filePicker)
    Else
        MsgBox "No file selected.", vbExclamation
        GoTo BeforeExit
    End If
    
    ' Loop through each sheet in ThisWorkbook (the workbook containing this code)
    For Each ws1 In ThisWorkbook.Sheets
        ws1.Columns("D:D").Insert Shift:=xlToRight
        Print #logFile, "WORKSHEET: " & ws1.Name & " - Changes"
        lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each cell in column A of the current sheet in ThisWorkbook
        For Each cell1 In ws1.Range("A1:A" & lastRow1)
            itemNumber = cell1.Value
            found = False
            
            ' Validate the item number
            If Len(itemNumber) = 6 And IsNumeric(itemNumber) Then
                ' Loop through each sheet in referenceBook to find "CI-DIRECT" column
                For Each ws2 In referenceBook.Sheets
                    headerFoundCount = 0
                    ' Check the first 25 rows for "CI-DIRECT" header
                    For Each headerCell In ws2.Range("A1:Z25") ' Assume the header is within columns A to Z
                        If Like headerCell.Value, "CI-DIRECT*" Then
                            headerFoundCount = headerFoundCount + 1
                            ciDirectCol = headerCell.Column
                            If headerFoundCount > 1 Then
                                MsgBox "Multiple 'CI-DIRECT' headers found in " & ws2.Name, vbCritical
                                GoTo BeforeExit
                            End If
                        End If
                    Next headerCell
                    
                    If headerFoundCount = 0 Then
                        MsgBox "'CI-DIRECT' header not found in " & ws2.Name, vbCritical
                        GoTo BeforeExit
                    End If
                    
                    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
                    ' Check for a matching item number in column A of the current sheet in referenceBook
                    Set r2 = ws2.Range("A1:A" & lastRow2).Find(What:=itemNumber, LookIn:=xlValues, LookAt:=xlWhole)
                    
                    ' If a match is found
                    If Not r2 Is Nothing Then
                        ' Update the cell in ThisWorkbook, newly added column D, with the value from "CI-DIRECT" column of referenceBook
                        ws1.Cells(cell1.Row, "D").Value = ws2.Cells(r2.Row, ciDirectCol).Value
                        found = True
                        
                        ' Log the update
                        Print #logFile, "Updated ThisWorkbook - " & ws1.Name & ", Cell: D" & cell1.Row & _
                                    " with value from referenceBook - " & ws2.Name & ", Cell: " & Cells(r2.Row, ciDirectCol).Address
                        
                        Exit For ' Exit after the first match
                    End If
                Next ws2
                
                ' If not found, write NOTFOUND
                If Not found Then
                    ws1.Cells(cell1.Row, "D").Value = "NOTFOUND"
                    Print #logFile, "NOTFOUND for ThisWorkbook - " & ws1.Name & ", Cell: A" & cell1.Row
                End If
            End If
        Next cell1
    Next ws1
    
    MsgBox "Operation Completed"
    
BeforeExit:
    ' Close the log file
    Close #logFile
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    
    ' Ensure the log file is closed in case of an error
    If logFile > 0 Then Close #logFile
End Sub
