Function MyFunction(param1 As Variant, param2 As Variant)
  ' Explicit Variant type declaration is recommended
  If IsNumeric(param1) And IsNumeric(param2) Then
    Result = param1 + param2
  Else
    Result = "Error: Parameters must be numeric"
  End If
  MyFunction = Result
End Function
'Alternatively, use On Error Resume Next to handle errors gracefully:
Function MyFunction2(param1 As Variant, param2 As Variant)
  On Error Resume Next
  Result = param1 + param2
  If Err.Number <> 0 Then
    Result = "Error: Parameters must be numeric"
    Err.Clear
  End If
  MyFunction2 = Result
End Function