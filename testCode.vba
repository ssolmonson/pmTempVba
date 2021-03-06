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
    Dim VarId() As Variant
    Dim UpComplete() As Variant
    Dim UpRemain() As Variant
    Dim UpStart() As Variant
    Dim UpFinish() As Variant
    Dim ColNotes() As Variant

    'Initiates starting index to use for the arrays
    Dim CountIndex As Integer
    CountIndex = 0

    'Determines Row
    Dim i As Integer
    'Starting row which contains data of value
    i = 8

'Loop through each row
    Do Until i = LastRowInWorksheet
        'Identify if column "I" has data, use an or statement to check subsequent columns
        'Column B is only used if there is value in the designated columns in the same row,
        'therefore it is not checked in the If statement
        If Range("I" & i).Value <> "" Or Range("K" & i).Value <> "" Or Range("O" & i).Value <> "" Or Range("Q" & i).Value <> "" Or Range("R" & i).Value <> "" Then
            'Add value to the designated B columns array at the index determined by the CountIndex variable
            ReDim Preserve VarId(CountIndex)
            VarId(CountIndex) = Range("B" & i).Value

            'Add value to the designated I columns array at the index determined by the CountIndex variable
            ReDim Preserve UpComplete(CountIndex)
            UpComplete(CountIndex) = Range("I" & i).Value

            'Add value to the designated K columns array at the index determined by the CountIndex variable
            ReDim Preserve UpRemain(CountIndex)
            UpRemain(CountIndex) = Range("K" & i).Value

            'Add value to the designated O columns array at the index determined by the CountIndex variable
            ReDim Preserve UpStart(CountIndex)
            UpStart(CountIndex) = Range("O" & i).Value

            'Add value to the designated Q columns array at the index determined by the CountIndex variable
            ReDim Preserve UpFinish(CountIndex)
            UpFinish(CountIndex) = Range("Q" & i).Value

            'Add value to the designated R columns array at the index determined by the CountIndex variable
            ReDim Preserve ColNotes(CountIndex)
            ColNotes(CountIndex) = Range("R" & i).Value

            'Add 1 to the counter, this will define indicies of the arrays
            CountIndex = CountIndex + 1

        End If

        'Add one to the i variable, indicating to move on to the next row
        i = i + 1

    Loop

    'Create Workbook
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

    'Create a new workbook in Excel
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add

    'Add headers to Worksheet
    Set oSheet = oBook.Worksheets(1)
    oSheet.Range("B1:G1").Value = Array("UID", "Updated %Complete", "Updated Remaining Cost/Work", "Updated Start", "Updated Start", "Updated Finish")

    'Transfer arrays to the worksheet
    'When using a 2D array use: oSheet.Range("").Resize(<rows>, <columns>).Value = DataArray

    'Loop through each array going down the specified column
    'Set variable to use for the index: y
    Dim y As Integer

    Do Until y = CountIndex

        oSheet.Range("B2").Offset(y, 0).Value = VarId(y)
        oSheet.Range("C2").Offset(y, 0).Value = UpComplete(y)
        oSheet.Range("D2").Offset(y, 0).Value = UpRemain(y)
        oSheet.Range("E2").Offset(y, 0).Value = UpStart(y)
        oSheet.Range("F2").Offset(y, 0).Value = UpFinish(y)
        oSheet.Range("G2").Offset(y, 0).Value = ColNotes(y)

        y = y + 1
    Loop



    'Save workbook and quit the newly created excel sheet
    oBook.SaveAs "C:\Users\Scott\Documents\TestReports\Report4.xlsx"
    oExcel.Quit

  'Test Code

    'Dim lastRow As Long
    'Dim lastCol As Long

        'lastRow = Cells(Rows.Count, "B").End(xlUp).Row

        'lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

        'MsgBox "Last Row: " & lastRow '& vbNewLine & _
                '" Last Column: " & lastCol

    'Finds the value of the cell
    'Worksheets("").Range("").Value

    'Used to test if the arrays are all the same size, so all data matches up to the correct UID
    'MsgBox ("The total length of all arrays is: " & CountIndex & " " & WorksheetFunction.CountA(VarId) & " " & WorksheetFunction.CountA(UpComplete) & " " & WorksheetFunction.CountA(UpRemain) & " " & WorksheetFunction.CountA(UpStart) & " " & WorksheetFunction.CountA(UpFinish) & " " & WorksheetFunction.CountA(ColNotes))

    'Test varying array indicies to ensure data is being entered correctly off a test file
    'MsgBox ("Test varying indicies: " & " " & VarId(2) & " " & VarId(6) & " " & Var(12) & " " & Var(24)

  'Debug.Print (lastRowInWorksheet)


End Sub
