Namespace ViewModel
 Public Class WorkspaceViewModel
  Inherits ViewModelBase

  Private _closeCommand As RelayCommand
  Public ReadOnly Property CloseCommand As RelayCommand
   Get
    If _closeCommand Is Nothing Then
     _closeCommand = New RelayCommand(Sub(param) Me.OnRequestClose())
    End If
    Return _closeCommand
   End Get
  End Property

  ''' <summary>
  ''' Raised when this workspace should be removed from the UI.
  ''' </summary>
  Public Event RequestClose As EventHandler
  Protected Overridable Sub OnRequestClose()
   RaiseEvent RequestClose(Me, EventArgs.Empty)
  End Sub
  Friend Sub CloseMe()
   RaiseEvent RequestClose(Me, EventArgs.Empty)
  End Sub

 End Class
End Namespace
