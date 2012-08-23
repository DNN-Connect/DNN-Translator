Public Class MainWindow

 Private Sub MainWindow_Initialized(sender As Object, e As System.EventArgs) Handles Me.Initialized
  Title = "DotNetNuke Translator " & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString
 End Sub

 'Private Sub cmdOptions_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdOptions.Click
 ' Dim sw As New SettingsWindow
 ' sw.DataContext = CType(DataContext, ViewModel.MainWindowViewModel).Settings
 ' sw.ShowDialog()
 'End Sub

End Class
