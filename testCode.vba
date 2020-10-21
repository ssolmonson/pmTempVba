Sub Test1()

'If cell is not empty in on yellow columns only
'If cell has an entry grab UID and put in new Excel sheet

'Hard code each column (in this case hard code columns (I, K, O, Q, R)

'End when sheet ends, find function to determine when spreadsheet is done, to avoid continuous loops down

'Define last used cells row in the range with End(xlUp)
    Dim LastRowInWorksheet As Long
    'End(xlUp) finds the last used cell in the range
    LastRowInWorksheet = Cells(Rows.Count, "B").End(xlUp).Row

    'Create Arrays for each column needed, including UID (Columns: B, I, K, O, Q, R)
    Dim UpComplete()
    Dim UpRemain()
    Dim UpStart()
    Dim UpFinish()
    Dim ColNotes()
    Dim VarId()

    Dim x As Integer
    Dim CountIndex As Integer
    Count = 0
    x = 8
'End if Range("I").End(xlUp).Select

'Loop through each row
    Do Until x = LastRowInWorksheet
        'Identify if column "I" has data, use an or statement to check subsequent columns
        If Range("I" & x).Value Or Range("K" & x).Value Or Range("O" & x).Value Or Range("Q" & x).Value Or Range("R" & x).Value Then
            'If Range("I" & x).Value <> "" Then
                'add value to I array at CountIndex
            'If "K" is not blank ""
                'add value to K array at CountIndex
            'If "O" is not blank ""
                'add value to O array at CountIndex
            'If "Q" is not blank ""
                'add value to Q array at CountIndex
            'If "R" is not blank ""
                'add value to R array at CountIndex
            'Add "B" UID to array

            'Add 1 to the counter, this will define indicies of the arrays

        'If IsEmpty(cell) is false
        'Find value of the cell, and te
        'Finds the value of the cell
        'Worksheets("").Range("").Value
        'If it has data grab the UID from column B and any data in the designated columns
        'Place data in Excel spreadsheet

    Next

'Exit when row is equal to row defined as one plus last row selected is that defined by End(xlUp)



'Test Code

    'Dim lastRow As Long
    'Dim lastCol As Long

        'lastRow = Cells(Rows.Count, "B").End(xlUp).Row

        'lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

        'MsgBox "Last Row: " & lastRow '& vbNewLine & _
                '" Last Column: " & lastCol

'Debug.Print (lastRowInWorksheet)

End Sub
