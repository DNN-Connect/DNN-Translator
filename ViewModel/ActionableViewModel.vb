Namespace ViewModel
 Public Class ActionableViewModel
  Inherits ViewModelBase

  Private _commandList As New List(Of CommandViewModel)

  Public Sub New(displayName As String)
   MyBase.DisplayName = displayName
   MyBase.ID = displayName
  End Sub

  Public Sub New(displayName As String, primaryCommandName As String, primaryCommand As ICommand)
   If Command() Is Nothing Then
    Throw New ArgumentNullException("command")
   End If
   MyBase.DisplayName = displayName
   MyBase.ID = displayName
   _commandList.Add(New CommandViewModel(primaryCommandName, primaryCommand))
  End Sub

  Public Sub AddCommand(displayName As String, command As ICommand)
   _commandList.Add(New CommandViewModel(displayName, command))
  End Sub
  Public Sub AddCommand(command As CommandViewModel)
   _commandList.Add(command)
  End Sub

  Public ReadOnly Property CommandList As List(Of CommandViewModel)
   Get
    Return _commandList
   End Get
  End Property

  Public ReadOnly Property Command As ICommand
   Get
    If _commandList.Count > 0 Then
     Return _commandList(0).Command
    Else
     Return Nothing
    End If
   End Get
  End Property

 End Class
End Namespace
