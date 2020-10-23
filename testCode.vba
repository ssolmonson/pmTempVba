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
    Dim VarId()
    Dim UpComplete()
    Dim UpRemain()
    Dim UpStart()
    Dim UpFinish()
    Dim ColNotes()

    'Determines Row
    Dim i As Integer
    'Initiates starting index to use for the arrays
    Dim CountIndex As Integer

    CountIndex = 0
    'Starting row which contains data of value
    i = 8

'Loop through each row
    Do Until i = LastRowInWorksheet
        'Identify if column "I" has data, use an or statement to check subsequent columns
        'Column B is only used if there is value in the designated columns in the same row,
        'therefore it is not checked in the If statement
        If Range("I" & i).Value Or Range("K" & i).Value Or Range("O" & i).Value Or Range("Q" & i).Value Or Range("R" & i).Value Then
            'Add value to the designated B columns array at the index determined by the CountIndex variable
            UpComplete(CountIndex) = Range("B" & i).Value

            'Add value to the designated I columns array at the index determined by the CountIndex variable
            UpComplete(CountIndex) = Range("I" & i).Value

            'Add value to the designated K columns array at the index determined by the CountIndex variable
            UpComplete(CountIndex) = Range("K" & i).Value

            'Add value to the designated O columns array at the index determined by the CountIndex variable
            UpComplete(CountIndex) = Range("O" & i).Value

            'Add value to the designated Q columns array at the index determined by the CountIndex variable
            UpComplete(CountIndex) = Range("Q" & i).Value

            'Add value to the designated R columns array at the index determined by the CountIndex variable
            UpComplete(CountIndex) = Range("R" & i).Value

            'Add 1 to the counter, this will define indicies of the arrays
            CountIndex = CountIndex + 1

        Next

        'Add one to the i variable, indicating to move on to the next row
        i = i + 1

    Loop

    MsgBox ("The total length of the arrays is: " & CountIndex)


'Test Code

    'Dim lastRow As Long
    'Dim lastCol As Long

        'lastRow = Cells(Rows.Count, "B").End(xlUp).Row

        'lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

        'MsgBox "Last Row: " & lastRow '& vbNewLine & _
                '" Last Column: " & lastCol

    'Finds the value of the cell
    'Worksheets("").Range("").Value

'Debug.Print (lastRowInWorksheet)

End Sub
