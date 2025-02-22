Function MyFunction(param1, param2)
  ' Missing type declaration for parameters
  If IsNumeric(param1) And IsNumeric(param2) Then
    Result = param1 + param2
  Else
    Result = "Error: Parameters must be numeric"
  End If
  MyFunction = Result
End Function