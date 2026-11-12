Attribute VB_Name = "SampleMigration"
Option Explicit

Public Sub ReconcileReport(ByVal inputPath As String, ByVal outputPath As String)
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    If Len(Trim$(inputPath)) = 0 Then Err.Raise vbObjectError + 100, "ReconcileReport", "Input path is required"
    Set sourceBook = Workbooks.Open(inputPath)
    Set sourceSheet = sourceBook.Worksheets("Input")
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 2 To lastRow
        If CStr(sourceSheet.Cells(rowIndex, 1).Value) = "" Then
            sourceSheet.Cells(rowIndex, 3).Value = "Missing key"
        Else
            ValidateRecord sourceSheet, rowIndex
        End If
    Next rowIndex

    sourceBook.SaveAs outputPath
    sourceBook.Close SaveChanges:=False
End Sub

Private Sub ValidateRecord(ByVal ws As Worksheet, ByVal rowIndex As Long)
    If CDbl(Val(ws.Cells(rowIndex, 2).Value)) <= 0 Then
        ws.Cells(rowIndex, 3).Value = "Invalid quantity"
    Else
        ws.Cells(rowIndex, 3).Value = "Validated"
    End If
End Sub
