Sub splitline()
' *************************************************************************************
' USER INPUT
Dim sheetn As String
Dim col As String
' UPDATE THIS VALUE FOR THE SHEET TO UPDATE
sheetn = "Sheet1"
' UPDATE THIS VALUE FOR THE COLUMN YOU WANT TO UPDATE ON
col = "B"
' *************************************************************************************


' *************************************************************************************
' HOW TO USE
' First, do a test run:
'   Copy a sheet to a new workbook
'   Use alt-f11 to open Macros
'   Add this code to any Sheet Object
'   Update the variables above
'   sheetn: the name of the sheet to update
'   col: the column to update on
'   Once these are updated, hit the green play button above
'   Verify that the output looks how you need

' Now, run on the sheets in question. Back up copies if worried
' about the macro imapcting anything

' *************************************************************************************

' MACRO VARIABLES
Dim rLastCell_1 As Range
Dim lLastRow_1  As Long
Dim genrowcnt As Integer

Dim RowCounter As Long
Dim TotalRows As Long
Dim rng As Range ' for getting next row & values
Dim CellValue As String ' value from the cell to parse for line breaks

Dim ArrLine As Variant
Dim ArrLen As Long
Dim Arr_i As Long
Dim Cur_Rng As Range

' Count Rows at Start
Set rLastCell_1 = Worksheets(sheetn).Cells.Find(What:="*", After:=Worksheets(sheetn).Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

If Not rLastCell_1 Is Nothing Then lLastRow_1 = rLastCell_1.Row


' Loop over rows
RowCounter = 1
TotalRows = lLastRow_1 + 1 ' updated during processing
genrowcnt = 0 ' updates as rows are added

Do While RowCounter < TotalRows
    ' get cell value
    Set rng = Worksheets(sheetn).Range(col & RowCounter)
    CellValue = rng
    
    ' parse for linebreaks, split to array
    ArrLines = Split(CellValue, vbLf)
    ArrLen = UBound(ArrLines) - LBound(ArrLines) + 1

    ' For each element in array, generate full copied new line, then keep only that value in cell
    If ArrLen > 1 Then
        ' Add n-1 rows below
        For Arr_i = LBound(ArrLines) To UBound(ArrLines) - 1
            ' MsgBox (ArrLines(Arr_i))
            ' Insert new row below current row (based on array position)
            rng.EntireRow.Offset(Arr_i + 1).Insert
            ' Copy this row and paste below
            Worksheets(sheetn).Rows(RowCounter).EntireRow.Copy Worksheets(sheetn).Range("A" & RowCounter + Arr_i + 1)
            genrowcnt = genrowcnt + 1
            TotalRows = TotalRows + 1
        Next Arr_i
        ' Update the cell values for current row + new created rows
        For Arr_i = LBound(ArrLines) To UBound(ArrLines)
            'Paste value at Arr_i into the Cell
            Set Cur_Rng = Worksheets(sheetn).Range(col & RowCounter + Arr_i)
            Cur_Rng.Value = ArrLines(Arr_i)
        Next Arr_i
    End If
    
    RowCounter = RowCounter + 1
Loop


' FINAL OUTPUT
MsgBox ("Generated " & genrowcnt & " new lines")


End Sub
