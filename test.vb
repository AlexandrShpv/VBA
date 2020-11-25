Sub test()
'
' test Macro
'

'
    Range("A2:I4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = -7235177
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("J9").Select
End Sub

Sub TO_ID_ramitis()
Dim ObjektaSakumaRinda As Integer
Dim RinduSkaits As Integer
Dim ObjektuRangeAddress As String
Dim ObjektuRange As Range

Application.Goto Range("A2")

ObjektaSakumaRinda = ActiveCell.Row
Debug.Print "Objekta sakuma rinda: " & ObjektaSakumaRinda
RinduSkaits = 0
    Do While ActiveCell.Value = ActiveCell.Offset(RinduSkaits, 0).Value
        RinduSkaits = RinduSkaits + 1
    Loop
Set ObjektuRange = Range("A" & ObjektaSakumaRinda & ":" & "I" & ObjektaSakumaRinda + RinduSkaits - 1)
Debug.Print "Objekta range adress: " & ObjektuRange.Address

with ObjektuRange.BorderAround

End Sub

