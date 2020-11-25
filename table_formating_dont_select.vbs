' Sean Johnson
' Excel VBA Range - Avoid Select, You will be glad you did
' https://www.youtube.com/watch?v=HaADHWl-UaA

Option Explicit
Sub GetReferenceToData()
    Dim rngData As Range
    Dim rngHeader As Range
    Dim rngColToFill As Range

    Set rngData = IhisWorkbook.Worksheets("SheetZ").Range("Al")
    Set rngData = rngData.CurrentRegion
    Set rngHeader = rngData.Resize(l)
    rngHeader.Font.Bold = True
    rngfieader.Font.Color = vaed
    Set rngColToFill = rngData.Resize(rngData.Rows.Count — 1, 2).Offset(l)
    rngColToFill.SpecialCells(xlCellTypeBlanks).FormulaRlCl = "=R[—l]C"
    rngColToFill.Value = rngColToFill.Value
End Sub