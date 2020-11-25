Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Range("A1"), Target) Is Nothing Then
    
    Application.ScreenUpdating = False
    
    Dim filterArray, c As Variant
    Dim filterArrayCount As Long
    Dim lcol As String
    Dim showFlag As Integer
    
    filterArray = Split(Cells(1, 1).Value, ", ")
    
        Dim filterRange As Range
        lcol = Rows(1).Find(What:="*", SearchDirection:=xlPrevious).Address
        Set filterRange = Range("B1:" & lcol)
        
        If Target.Value = "" Then
            For Each c In filterRange
                ActiveSheet.Range(c.Address).EntireColumn.Hidden = False
            Next c
        Else
            For Each c In filterRange
                showFlag = 0
                For filterArrayCount = 0 To UBound(filterArray)
                    If Left(filterArray(filterArrayCount), 1) <> "-" Then
                        If InStr(UCase(c.Value), UCase(filterArray(filterArrayCount))) > 0 Then
                            showFlag = showFlag + 1
                        End If
                    Else
                        If InStr(UCase(c.Value), UCase(Right(filterArray(filterArrayCount), Len(filterArray(filterArrayCount)) - 1))) > 0 Then
                            showFlag = 0
                        End If
                    End If
                Next filterArrayCount
                If showFlag > 0 Then
                    ActiveSheet.Range(c.Address).EntireColumn.Hidden = False
                Else
                    ActiveSheet.Range(c.Address).EntireColumn.Hidden = True
                End If
            Next c
        End If
        
        
        Application.ScreenUpdating = True
        
    End If

End Sub
