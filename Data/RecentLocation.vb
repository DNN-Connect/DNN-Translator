Namespace Data
 Public Class RecentLocation
  Public Property Location As String = ""
  Public Property LocationShort As String = ""
  Public Property Index As Integer = -1

  Public Sub New(index As Integer, location As String)
   _Index = index
   _Location = location
   If location.Length > 50 Then
    LocationShort = location.Substring(0, 3) & "..." & location.Substring(location.Length - 40, 40)
   Else
    LocationShort = location
   End If
  End Sub
 End Class
End Namespace