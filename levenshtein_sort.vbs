Private Sub Worksheet_Change(ByVal Target As Range)
  
  Dim lo As ListObject
  Dim sortRange As Range

  Set lo = ActiveWorkbook.Sheets("Konfiguracija_CSV").ListObjects("tabKonfTipi")
  Set sortRange = ActiveWorkbook.Sheets("Konfiguracija_CSV").ListObjects("tabKonTipi").ListColumns("Levenshtein").Range

  If Not Application.Intersect(Range("E2"), Target) Is Nothing Then
    Application.Volatile ' Wait Automatic calculation finish
    With lo.sort
      .SortFields.Clear
      .SortFields.Add Key:=sortRange, Order:=xlAscending
      .Header = xlYes
      .Apply
    End With
  End If

  If Not Application.Intersect(Range("E1"), Target) Is Nothing Then
    With lo.Range
      .Autofilter Field:=5, Criteria1:="*" & Cells(1,5).Value & "*"
    End With
  End If
End Sub