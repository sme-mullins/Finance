Attribute VB_Name = "Module1"
Sub DistributeData()
    On Error GoTo ErrorHandler

    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim sourceRange As Range
    Dim lastRow As Long
    
    ' Set the source worksheet (adjust the sheet name as needed)
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row with data in the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' Select all fields from the source sheet
    Set sourceRange = wsSource.Range("A1:D" & lastRow)
    
    ' Create a new sheet titled "Journal Entry"
    Set wsDestination = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsDestination.Name = "Journal Entry"
    
    ' Define the columns in the new sheet
    Dim headers As Variant
    headers = Array("JOURNAL", "DATE", "DESCRIPTION", "SOURCEENTITY", "LINE_NO", "ACCT_NO", "LOCATION_ID", "DEPT_ID", "GLENTRY_CLASSID", "DEBIT", "CREDIT", "MEMO", "STATE")
                    
    Dim i As Long ' Change to Long for consistency
    For i = LBound(headers) To UBound(headers)
        wsDestination.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ' Copy data from source sheet to the new sheet
    Dim srcRow As Long
    Dim destRow As Long
    destRow = 2 ' Start copying to the second row of the destination sheet
    
    For srcRow = 1 To lastRow
        With wsDestination
            Dim sourceValue As Variant
            sourceValue = wsSource.Cells(srcRow, 1).Value
            
            ' Assign constant values to specified columns
            .Cells(destRow, 1).Value = "GJ" ' JOURNAL
            .Cells(destRow, 4).Value = 1 ' SOURCEENTITY
            .Cells(destRow, 13).Value = "Draft" ' STATE
            .Cells(destRow, 9).Value = 1 ' GLENTRY_CLASSID
            
            ' Assign MEMO from Sheet1 Column C to MEMO column in journal entry sheet
            .Cells(destRow, 12).Value = wsSource.Cells(srcRow, 3).Value ' MEMO
            
            ' Extract and assign values from Column A
            .Cells(destRow, 6).Value = Mid(sourceValue, 9, 4) ' ACCT_NO
            .Cells(destRow, 7).Value = Mid(sourceValue, 1, 1) ' LOCATION_ID
            .Cells(destRow, 14).Value = Mid(sourceValue, 15, 1) ' SUBLOCATION_ID
            
            ' Check and adjust DEPT_ID
            Dim deptID As String
            deptID = Mid(sourceValue, 5, 3) ' Extract DEPT_ID from Column A
            
            If Val(deptID) <= 399 Then
                deptID = "100"
            End If
            
            .Cells(destRow, 8).Value = deptID ' Assign adjusted DEPT_ID
            
            ' Assign the second column from Sheet1 to the "DEBIT" column in the journal entry sheet
            .Cells(destRow, 10).Value = wsSource.Cells(srcRow, 2).Value ' DEBIT
            
            ' Assign the fourth column from Sheet1 to the "DATE" column in the journal entry sheet
            .Cells(destRow, 2).Value = wsSource.Cells(srcRow, 4).Value ' DATE
                   
            ' Check if GL account is 4440, department is 514, and sublocation is 1
            Dim glAccount As String
            glAccount = Mid(sourceValue, 9, 4) ' Extract GL_ACCOUNT from Column A
            
            If glAccount = "4440" And deptID = "514" And .Cells(destRow, 14).Value = "1" Then
                glAccount = "4443"
            End If
            
            .Cells(destRow, 6).Value = glAccount ' Assign adjusted ACCT_NO from GL_ACCOUNT
            
            ' Assign DESCRIPTION as 'Integration' + DATE
            .Cells(destRow, 3).Value = "Integration " & Format(wsSource.Cells(srcRow, 4).Value, "mm/dd/yyyy")
        End With
        destRow = destRow + 1
    Next srcRow
    
    ' Format date column
    wsDestination.Range("B2:B" & lastRow).NumberFormat = "mm/dd/yyyy"
    
    ' Call the CountUniqueDates subroutine to get the number of unique dates
    CountUniqueDates wsDestination
    
    ' Notify the user that the process is complete
    MsgBox "Data has been successfully imported and disseminated into the 'Journal Entry' sheet.", vbInformation
    
Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub

Sub CountUniqueDates(wsDestination As Worksheet)
    On Error Resume Next ' Skip errors in case of non-date values

    Dim lastRow As Long
    Dim dateCell As Range
    Dim dateCountDict As Object
    Dim dateValue As Variant
    
    ' Find the last row with data in the DATE column
    lastRow = wsDestination.Cells(wsDestination.Rows.count, 3).End(xlUp).Row
    
    ' Create a dictionary to store date counts
    Set dateCountDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the DATE column to count occurrences of each date and assign counts to LINE_NO column
    For Each dateCell In wsDestination.Range("B2:B" & lastRow)
        If IsDate(dateCell.Value) Then
            dateValue = CStr(dateCell.Value)
            If Not dateCountDict.Exists(dateValue) Then
                dateCountDict.Add dateValue, 1
            Else
                dateCountDict(dateValue) = dateCountDict(dateValue) + 1
            End If
            dateCell.Offset(0, 3).Value = dateCountDict(dateValue) ' Assign the count to LINE_NO column (3 columns right of DATE)
        End If
    Next dateCell
    
    On Error GoTo 0 ' Restore default error handling
End Sub

