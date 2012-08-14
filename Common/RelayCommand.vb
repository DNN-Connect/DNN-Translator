' Thanks to http://waxtadpole.wordpress.com/2009/10/21/wpf-mvvm-in-vb-net-delegating-event-subscriptions-for-routedcommand-or-any-icommand-in-the-viewmodel/

Public Class RelayCommand
 Implements ICommand

#Region " Fields "
 Private _execute As Action(Of Object)
 Private _canExecute As Predicate(Of Object)
#End Region

#Region " Constructors "
 Public Sub New(execute As Action(Of Object))
  Me.New(execute, Nothing)
 End Sub

 Public Sub New(execute As Action(Of Object), canExecute As Predicate(Of Object))
  If execute Is Nothing Then
   Throw New ArgumentNullException("execute")
  End If
  _execute = execute
  _canExecute = canExecute
 End Sub
#End Region

#Region " ICommand Members "

 <DebuggerStepThrough()> _
 Public Function CanExecute(parameter As Object) As Boolean Implements System.Windows.Input.ICommand.CanExecute
  Return If(_canExecute Is Nothing, True, _canExecute(parameter))
 End Function

 Public Custom Event CanExecuteChanged As EventHandler Implements System.Windows.Input.ICommand.CanExecuteChanged
  AddHandler(ByVal value As EventHandler)
   AddHandler CommandManager.RequerySuggested, value
  End AddHandler
  RemoveHandler(ByVal value As EventHandler)
   RemoveHandler CommandManager.RequerySuggested, value
  End RemoveHandler
  RaiseEvent(ByVal sender As Object, ByVal e As System.EventArgs)
   CommandManager.InvalidateRequerySuggested()
  End RaiseEvent
 End Event

 Public Sub Execute(parameter As Object) Implements System.Windows.Input.ICommand.Execute
  _execute(parameter)
 End Sub

#End Region

End Class
