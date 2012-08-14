Imports System.Windows
Imports System.Windows.Input
Imports DotNetNuke.Translator.Controls.FolderBrowser

Namespace Controls
 Public Class FolderBrowserDialog
  Inherits Window
  Private _viewModel As BrowserViewModel

  Public ReadOnly Property ViewModel() As BrowserViewModel
   Get
    Return InlineAssignHelper(_viewModel, If(_viewModel, New BrowserViewModel()))
   End Get
  End Property

  Public Property SelectedFolder() As String
   Get
    Return ViewModel.SelectedFolder
   End Get
   Set(value As String)
    ViewModel.SelectedFolder = value
   End Set
  End Property

  Private Sub Ok_Click(sender As Object, e As RoutedEventArgs)
   DialogResult = True
  End Sub

  Private Sub TextBlock_MouseDown(sender As Object, e As MouseButtonEventArgs)
   If e.ClickCount = 2 AndAlso e.LeftButton = MouseButtonState.Pressed Then
    ' close the dialog on a double-click of a folder
    DialogResult = True
   End If
  End Sub
  Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
   target = value
   Return value
  End Function

 End Class
End Namespace
