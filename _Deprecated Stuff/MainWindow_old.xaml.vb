Imports System.Collections.ObjectModel
Imports DotNetNuke.Translator.Common

Class MainWindow
 Inherits RibbonWindow

 Private Sub MainWindow_Initialized(sender As Object, e As System.EventArgs) Handles Me.Initialized
  Title = "DotNetNuke Translator " & System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString
  AddHandler ViewModel.TranslatorModel.OpenNewLocation, AddressOf OpenNewLocation
 End Sub

 Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
  ViewModel.TranslatorModel.ExitApplication()
 End Sub

 Private Sub RecentLocationSelect(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs)
  Dim tb As TextBlock = CType(sender, TextBlock)
  Dim selectedLocation As String = CType(tb.Tag, String)
  OpenNewLocation(selectedLocation)
 End Sub

 Private Sub OpenNewLocation(location As String)

  'ContentPanel.Children.Clear()


  'Dim p As New ProgressBar
  'ContentPanel.Children.Add(p)
  'ContentPanel.Refresh()
  'p.Height = 12
  'Dim binding As New Binding
  'binding.Source = ContentPanel
  'binding.Path = New PropertyPath("ActualWidth")
  'p.SetBinding(FrameworkElement.WidthProperty, binding)
  'p.Margin = New Thickness(6)
  'p.Minimum = 0
  'p.Maximum = l.Length + 1
  'p.Value = 0

  Dim d As New IO.DirectoryInfo(location)
  ViewModel.TranslatorModel.OpenedNewLocation(location)



 End Sub

End Class
