Attribute VB_Name = "CombineSheets_SendTasks"
Option Explicit
'' Create by Noam Brand 23/3/2022
'' Purpose: Project assignment: Automatically create personalized tables of assignments for each employee, allowing them
'' to focus on only their assignments and priorities and not be distracted by other employee assignments that do not concern them.
'' Create managment reports with all projects and employees together.

'' The example has 4 sheets- 3 big projects that have many assignments and 1 sheet that has 4 small projects with few assignments.
'' Ps. If the assignment is e.g. "Jerry+Noam" it will be present in Jerry's table and also Noam's table.

'' Question: Why not create a table with all the projects and assignments in the first place that will save the need to combine different sheets?
'' Answer: When your table gets filled with projects that have many assignments it becomes difficult to manage efficiently, the table gets messy and
'' editing is more complicated with autofilter on,  you have to keep filtering the table which is time-consuming and prone to errors.

'' How does it work: It combines different sheets with the same fields to one table, copies the combined table to a new sheet and
'' Filters it by employee name to separate sheets, AutoFits the table nicely and finally saves the sheets to separate excel files
'' in a new folder that has the name of the workbook.
'' The AutoFilter in the code is in columns E,F (See Fields 5,6 in the AutoFilter)

'' Macro steps:
''1) Combines Number of first sheets with the same fields to one sheet.
''2) Makes copies the combined sheet to the end of the workbook.
''3) Rename sheets by each employee name.
''4) Filter each sheet by employee name +not yet complite assignments.
''5) Each sheet still has all of the assignments(the user can cancel filter by autofilter).
''6) Autofit columns width and sheet from right to left
''7) Split the workbook sheets to separate files in a folder.
''   The file name will be the sheet name + date of creation.

Sub main()
Dim iLastRow As Long
iLastRow = sheetfilter.Range("a999").End(xlUp).Row
Dim CountsheetsToCombine As Long
CountsheetsToCombine = sheetfilter.Cells(3, 3).Value

With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = False
      .DisplayAlerts = False
End With
  
  
CompareRows
Combine
On Error Resume Next
''Rename sheets and AutoFit
Dim projectRow As Long
For projectRow = 3 To iLastRow - 1
        CopySheetToEnd
        ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets.Count).Activate 'Activate last sheet
        ActiveWorkbook.ActiveSheet.Name = sheetfilter.Cells(projectRow, 1).Value
        FilterRows (projectRow)
        AutoFitColumns
Next
RTLsheet
FreezeFirstRow
''''''''''''''''''''''''''''''
' "all projects sheet"- the sheet that is combined without filter
ActiveWorkbook.Sheets(1).Activate
ActiveWorkbook.Sheets(1).Name = sheetfilter.Cells(iLastRow, 1).Value 'all projects
ActiveWorkbook.Sheets(1).Move after:=ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets.Count)  'move to end
AutoFitColumns
''''''''''''''''''''''''''''''

With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = True
      .DisplayAlerts = True
End With

Exit Sub
skip:

End Sub

''https://www.extendoffice.com/documents/excel/1184-excel-merge-multiple-worksheets-into-one.html
Sub Combine()
Dim J As Long
Dim CountsheetsToCombine As Long
On Error Resume Next
ActiveWorkbook.Sheets(1).Select
ActiveWorkbook.Worksheets.Add
ActiveWorkbook.Sheets(1).Name = "Combined"
ActiveWorkbook.Sheets(2).Activate
ActiveSheet.Range("A1").EntireRow.Select
Selection.Copy Destination:=ActiveWorkbook.Sheets(1).Range("A1")
CountsheetsToCombine = sheetfilter.Cells(3, 3).Value

For J = 2 To 2 + CountsheetsToCombine - 1
    ActiveWorkbook.Sheets(J).Activate
    ActiveSheet.Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
    Selection.Copy Destination:=ActiveWorkbook.Sheets(1).Range("A9999").End(xlUp)(2)
Next
End Sub


Sub FilterRows(ByVal projectRow As Long)
On Error Resume Next
'If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
Dim ARY(1)
With ActiveWorkbook.ActiveSheet
    .Range("A1:L500").AutoFilter Field:=6  'clear filter in Field=6
    .Range("A1:L500").AutoFilter Field:=5  'clear filter in Field=5
    ARY(0) = sheetfilter.Cells(projectRow, 1).Value  '' "*איגור*"'
    ARY(1) = sheetfilter.Cells(3, 2).Value ''="<>בוצע"
    .Range("A1:L500").AutoFilter Field:=6, Criteria1:="*" & ARY(0) & "*", Operator:=xlFilterValues
    .Range("A1:L500").AutoFilter Field:=5, Criteria1:="<>" & ARY(1), Operator:=xlFilterValues
End With
End Sub



Public Sub CopySheetToEnd()
On Error GoTo skip
    ActiveWorkbook.Sheets(1).Copy after:=ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets.Count)
Exit Sub
skip:
 Stop
End Sub

Sub FreezeFirstRow()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
Next
End Sub

    
Sub RTLsheet()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.DisplayRightToLeft = True
    Next
End Sub


Sub AutoFitColumns()
ActiveWorkbook.ActiveSheet.Columns("A:L").AutoFit
ActiveWorkbook.ActiveSheet.Columns("B").ColumnWidth = 5
ActiveWorkbook.ActiveSheet.Columns("D").ColumnWidth = 50
ActiveWorkbook.ActiveSheet.Rows("1:300").AutoFit
End Sub


Sub SplitWorkbook()
Dim FileExtStr As String
Dim xFile As String
Dim FileFormatNum As Long
Dim xWs As Worksheet
Dim xWb As Workbook
Dim xNWb As Workbook
Dim FolderName As String
Dim DateString As String

With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = False
      .DisplayAlerts = False
End With

Set xWb = Application.ThisWorkbook

DateString = Format$(Now, "dd-mm-yyyy")
FolderName = xWb.Path & "\" & DateString & " " & xWb.Name

If Val(Application.Version) < 12 Then
    FileExtStr = ".xls": FileFormatNum = -4143
Else
    Select Case xWb.FileFormat
        Case 51:
            FileExtStr = ".xlsx": FileFormatNum = 51
        Case 52:
            If Application.ActiveWorkbook.HasVBProject Then
                FileExtStr = ".xlsm": FileFormatNum = 52
            Else
                FileExtStr = ".xlsx": FileFormatNum = 51
            End If
        Case 56:
            FileExtStr = ".xls": FileFormatNum = 56
        Case Else:
            FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
End If
Debug.Print FolderName
MkDir FolderName

For Each xWs In xWb.Worksheets
On Error GoTo NErro
    If xWs.Visible = xlSheetVisible Then
    xWs.Select
    xWs.Copy
    xFile = FolderName & "\" & xWs.Name & " " & DateString & FileExtStr
    Set xNWb = Application.Workbooks.Item(Application.Workbooks.Count)
    xNWb.SaveAs xFile, FileFormat:=FileFormatNum
    xNWb.Close False, xFile
    End If
NErro:
    xWb.Activate
Next


With Application
      .Calculation = xlCalculationManual
      .ScreenUpdating = True
      .DisplayAlerts = True
End With

    MsgBox "The files are saved in: " & vbCrLf & FolderName
End Sub



Sub DeleteDuplicateSheets()
'https://powerspreadsheets.com/excel-vba-delete-sheet/
On Error Resume Next
Dim projectRow As Long
Dim ws As Worksheet
Dim iLastRow As Long
iLastRow = sheetfilter.Range("a999").End(xlUp).Row

Application.DisplayAlerts = False

For Each ws In ActiveWorkbook.Worksheets
Debug.Print ws.Name
   For projectRow = 3 To iLastRow + 1
        If ws.Name = sheetfilter.Cells(projectRow, 1).Value Then
           ws.Delete
        End If
    Next
Next ws

Application.DisplayAlerts = True
End Sub

' Check that all projects sheets have similar fields in first row
Function CompareRows() As Boolean
Dim J As Long
Dim ws As Worksheet
Dim CountsheetsToCombine As Long
Dim rng1 As Range, rng2 As Range
Dim mycolrng1 As Long, mycolrng2 As Long
Dim res As Boolean
CountsheetsToCombine = sheetfilter.Cells(3, 3).Value

With ActiveWorkbook.Sheets(1)
    mycolrng1 = .Range("A1").End(xlToRight).Column
    Set rng1 = Range(.Cells(1, 1), .Cells(1, mycolrng1))
End With

''compare all sheets
For J = 2 To CountsheetsToCombine
    With ActiveWorkbook.Sheets(J)
        mycolrng2 = .Range("A1").End(xlToRight).Column
        Set rng2 = .Range(.Cells(1, 1), .Cells(1, mycolrng2))
    End With
    CompareRows = RowsSimilar(rng1, rng2, J)
    If CompareRows = False Then
        MsgBox rng2.Worksheet.Name & " and" & rng1.Worksheet.Name & " fields are not diffrent"
        CampreSheets.Activate
        End
    End If
Next
End Function


Function RowsSimilar(r1 As Range, r2 As Range, J As Long) As Boolean
Dim i As Long

        For i = 1 To r1.Columns.Count
            If Not StrComp(r1.Cells(1, i), r2.Cells(1, i), vbBinaryCompare) = 0 Then
                    With ActiveWorkbook.Sheets(J)
                    RowsSimilar = False
'                         .Range(r2.Cells(1, i), r2.Cells(1, i)).Interior.ColorIndex = 6
                    Exit Function
                    End With
            End If
        Next i
RowsSimilar = True
End Function
