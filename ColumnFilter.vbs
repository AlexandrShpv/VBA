Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Range("A1"), Target) Is Nothing Then
    
    Application.ScreenUpdating = False
    
    'Filter Array
    Dim fA, c As Variant
    'Filter Array Counter
    Dim fAC As Long
    'Last head range column (also hidden)
    Dim lC As String
    ' Flag to remain column visible
    Dim sF As Integer
    
    fA = Split(Cells(1, 1).Value, ", ")
    
        Dim filterRange As Range
        lC = Rows(1).Find(What:="*", SearchDirection:=xlPrevious).Address
        Set filterRange = Range("B1:" & lC)
        
        If Target.Value = "" Then
            For Each c In filterRange
                ActiveSheet.Range(c.Address).EntireColumn.Hidden = False
            Next c
        Else
            For Each c In filterRange
                sF = 0
                For fAC = 0 To UBound(fA)
                    If Left(fA(fAC), 1) <> "-" Then
                        If InStr(UCase(c.Value), UCase(fA(fAC))) > 0 Then
                            sF = sF + 1
                        End If
                    Else
                        If InStr(UCase(c.Value), UCase(Right(fA(fAC), Len(fA(fAC)) - 1))) > 0 Then
                            sF = 0
                        End If
                    End If
                Next fAC
                If sF > 0 Then
                    ActiveSheet.Range(c.Address).EntireColumn.Hidden = False
                Else
                    ActiveSheet.Range(c.Address).EntireColumn.Hidden = True
                End If
            Next c
        End If
        
        
        Application.ScreenUpdating = True
        ActiveWindow.ScrollColumn = 1
        
    End If

End Sub
