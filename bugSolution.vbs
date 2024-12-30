Function MyFunc(param)
  On Error Resume Next
  If IsEmpty(param) Then
    Err.Raise 9999, , "Param cannot be empty"
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbCritical
        Err.Clear
        Exit Function ' or return a default value
    End If
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function