Namespace Common
 Public NotInheritable Class ApplicationState
  Private Shared _values As New Dictionary(Of String, Object)()
  Public Shared Sub SetValue(key As String, value As Object)
   If _values.ContainsKey(key) Then
    _values.Remove(key)
   End If
   _values.Add(key, value)
  End Sub

  Public Shared Function GetValue(Of T)(key As String) As T
   If _values.ContainsKey(key) Then
    Return DirectCast(_values(key), T)
   Else
    Return Nothing
   End If
  End Function
 End Class
End Namespace