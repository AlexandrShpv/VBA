' Range to String (error)

Function Concat(myRange as Range)
Dim myStr As String
Dim c As Range
myStr = ""
For Each c In myRange
  If c.Value <> "" Then myStr = myStr & ", " & c.Value
Next
Concat = Mid(myStr, 2, 9999)
End Function

