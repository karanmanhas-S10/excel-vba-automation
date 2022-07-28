
Sub DateRangerFilter()
''Variable Setting
Dim wb As Workbook
Dim ws As Worksheet
Dim Rng As Range
Dim CRng As Range ' Criteria Range = CR
Dim DRng As Range ' Destination Range = DR

Set wb = ThisWorkbook
Set ws = wb.Worksheets("Data Process")
Set ws2 = wb.Worksheets("Output")
Set Rng = ws.Range("B9").CurrentRegion
Set CRng = ws.Range("U2").CurrentRegion
Set DRng = ws2.Range("B10")

DRng.CurrentRegion.Clear

Rng.AdvancedFilter xlFilterCopy, CRng, DRng






End Sub
