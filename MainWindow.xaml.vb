Public Class MainWindow

 Private Sub MainWindow_Initialized(sender As Object, e As System.EventArgs) Handles Me.Initialized
  Title = "DotNetNuke Translator " & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString
  Me.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
 End Sub

 Private Sub MainWindow_SourceInitialized(sender As Object, e As System.EventArgs) Handles Me.SourceInitialized
  Me.WindowState = Windows.WindowState.Maximized
 End Sub
End Class
