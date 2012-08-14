Namespace ViewModel
 Public Class CommandViewModel
  Inherits ViewModelBase

  Public Property Command As ICommand
  Public Property CommandParameter As Object

  Public Sub New(displayName As String, command As ICommand)
   If command Is Nothing Then
    Throw New ArgumentNullException("command")
   End If
   MyBase.DisplayName = displayName
   MyBase.ID = displayName
   Me.Command = command
  End Sub

  Public Sub New(displayName As String, command As ICommand, commandParameter As Object)
   If command Is Nothing Then
    Throw New ArgumentNullException("command")
   End If
   MyBase.DisplayName = displayName
   MyBase.ID = displayName
   Me.Command = command
   Me.CommandParameter = commandParameter
  End Sub

 End Class
End Namespace