Function f(x)
  If x < 0 Then
    f = -1
  ElseIf x = 0 Then
    f = 0
  Else
    f = 1
  End If
End Function
MsgBox f(-1)
'The above code will return 0

'Corrected Code:
Function f(x)
  Dim result As Integer 'Explicit type declaration for result
  If x < 0 Then
    result = -1
  ElseIf x = 0 Then
    result = 0
  Else
    result = 1
  End If
  f = result 'Assign the result to the function's return value
End Function
MsgBox f(-1)