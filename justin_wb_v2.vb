'This is intended to take the file generated from the CRM and split it evenly among the number of reps
'declared by the user group. These files can then be uploaded as source lists to the various CRM
'platforms for leads.
'Requested: Justin Meldrum jmeldrum@jumpcrew.com
'Author: James Collier jcollier@jumpcrew.com
'Date: 01Jun2021
'
Option Explicit

Public SheetNumInput As Integer
Public Distribution As Long
    
Public Sub MyInputBox()

    Dim myBook As Workbook
    Dim i As Integer
    Dim NewSheet As Worksheet
    
    SheetNumInput = InputBox("Input number of desired tabs", "Rep Count", "Fill in integer")

    Worksheets("Controls").Activate
        Range("A2") = SheetNumInput
        
    Set myBook = ActiveWorkbook
    
For i = 1 To SheetNumInput
    With myBook
        'New sheet
        Set NewSheet = .Worksheets.Add(After:=.Worksheets("CSV"))
    End With

    NewSheet.Name = "Rep" & i

Next i

Worksheets("Controls").Activate

End Sub

Public Sub Import_CSV()
    Dim wb_CSV As Workbook
    Dim wb_Report As Workbook
    Set wb_Report = ActiveWorkbook
    Dim csvPath As String
    Dim Last_Row_CSV As Integer

    Dim FileToOpen As String
    
    FileToOpen = Application.GetOpenFilename
    
    Worksheets("Controls").Activate
    csvPath = FileToOpen
    Set wb_CSV = Workbooks.Open(csvPath)

    Last_Row_CSV = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    wb_CSV.Sheets(1).Range("A1:FT" & Last_Row_CSV).Copy
    wb_Report.Sheets("CSV").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    wb_CSV.Close
    Set wb_CSV = Nothing
    
    Worksheets("Controls").Activate
End Sub

Public Sub Clean_WB()
    Dim myBook As Workbook
    Dim ws As Worksheet
    Dim j As Integer
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("CSV").Delete
    
    'delete old rep tabs
    For j = 1 To SheetNumInput
        ThisWorkbook.Sheets("Rep" & j).Delete
    Next j
    
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:= _
                 ThisWorkbook.Sheets("Controls"))
        ws.Name = "CSV"
        
    Worksheets("Controls").Activate
End Sub

Public Sub CountCsvRows()
    Dim last_row As Long
    Dim denom As Integer
    
    Worksheets("CSV").Activate
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    Distribution = last_row / SheetNumInput
    
    Worksheets("Controls").Activate
    Range("A4") = Distribution
    
End Sub

Public Sub SplitPopulation()
'Update range to copy based on columns of the CSV file being imported
    Dim rng_start As Integer
    Dim rng_end As Integer
    Dim k As Integer
    Dim j As Integer
    
   'copy header row of csv
   For j = 1 To SheetNumInput
        Worksheets("Rep" & j).Activate
        Sheets("CSV").Range("A1:FT1").Copy
        Sheets("Rep" & j).Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   Next j
   
   'copy body of csv
    rng_start = 2
    rng_end = Distribution
    For k = 1 To SheetNumInput
        Worksheets("Rep" & k).Activate
        Sheets("CSV").Range("A" & rng_start & ":FT" & rng_end).Copy
        Sheets("Rep" & k).Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        rng_start = rng_end + 1
        rng_end = rng_end + Distribution + 1
    Next k
End Sub
