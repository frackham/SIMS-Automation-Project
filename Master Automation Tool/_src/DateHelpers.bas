Attribute VB_Name = "DateHelpers"

Option Explicit


Public Sub TimeStampThisLastModified_Update()
    'Sets the last modifed date of this workbook.
    ThisWorkbook.Worksheets("Set up").Range("J31").Value = Now
End Sub

Public Function TimeStampThisLastModified() As Date
    'Sets the last modifed date of this workbook.
    TimeStampThisLastModified = ThisWorkbook.Worksheets("Set up").Range("J31").Value
End Function

Public Function GetIsCheckDateBeforeDate(dDateToCheck As Date, dComparisonDate As Date) As Boolean
    GetIsCheckDateBeforeDate = False
    If dDateToCheck < dComparisonDate Then GetIsCheckDateBeforeDate = True
End Function




Public Function TimeStampNow() As Date
    TimeStampNow = Now
End Function
