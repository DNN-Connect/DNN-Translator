Namespace Common
 Public Class TreeItem

  Public Property Name As String = ""
  Public Property FullName As String = ""
  Public Property Children As New List(Of TreeItem)

  Public Sub New(name As String, fullName As String)
   Me.Name = name
   Me.FullName = fullName
  End Sub

  Public Function HasChild(name As String) As Boolean
   For Each child As TreeItem In Me.Children
    If child.Name = name Then Return True
   Next
   Return False
  End Function

  Public Function GetChild(name As String) As TreeItem
   For Each child As TreeItem In Me.Children
    If child.Name = name Then Return child
   Next
   Return Nothing
  End Function

  Public Function AddChild(name As String) As TreeItem
   For Each child As TreeItem In Me.Children
    If child.Name = name Then Return child
   Next
   Dim path As String = FullName
   If path <> "" Then path &= "\"
   Dim newChild As New TreeItem(name, path & name)
   Me.Children.Add(newChild)
   Return newChild
  End Function

  Public Function FindCommonEntryPoint() As TreeItem

   If Me.Children.Count = 0 Then Return Nothing
   If Me.Children.Count > 1 Then Return Me
   If Me.Children(0).Name.EndsWith(".resx") Then Return Me
   Return Me.Children(0).FindCommonEntryPoint

  End Function

 End Class
End Namespace