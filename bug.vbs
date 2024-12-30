Function MyFunc(param)
  If IsEmpty(param) Then
    Err.Raise 9999, , "Param cannot be empty"
  End If
  ' ... rest of the function
End Function