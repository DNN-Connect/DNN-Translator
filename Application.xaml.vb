Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.
 Shared dc As New Common.DataGridContextHelper()

 Protected Overrides Sub OnStartup(e As System.Windows.StartupEventArgs)
  MyBase.OnStartup(e)

  Dim window As New MainWindow

  Dim viewModel As New Translator.ViewModel.MainWindowViewModel
  AddHandler viewModel.RequestClose, Sub(sender As Object, e2 As System.EventArgs) window.Close()

  window.DataContext = viewModel

  window.Show()

 End Sub

End Class
