Imports DotNetNuke.Translator.ViewModel

Namespace Common
 Public Class ParameterObject

  Public Params As List(Of String)
  Public ParentWindow As ViewModelBase

  Public Sub New(parentWindow As ViewModelBase)
   Me.ParentWindow = parentWindow
   Me.Params = New List(Of String)
  End Sub

 End Class
End Namespace
