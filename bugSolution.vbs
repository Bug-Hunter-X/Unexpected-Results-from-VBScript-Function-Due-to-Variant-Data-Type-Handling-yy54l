Function CorrectCalculation(param1, param2)
  'Explicitly define data types for better type handling
  Dim result As Double

  If IsNumeric(param1) And IsNumeric(param2) Then
    result = CDbl(param1) * CDbl(param2) 'Explicit type conversion
  Else
    result = 0 'Handle non-numeric inputs appropriately
  End If

  CorrectCalculation = result
End Function