Namespace ViewModel
 Public Class BrowserViewModel
  Inherits WorkspaceViewModel

  Public Property Url As String

  Public Sub New(url As String)
   Me.Url = url
   Me.OnPropertyChanged("Url")
  End Sub

 End Class
End Namespace
