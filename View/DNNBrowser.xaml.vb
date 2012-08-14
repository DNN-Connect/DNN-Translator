Imports DotNetNuke.Translator.ViewModel

Namespace View
 Public Class DNNBrowser

  Public Sub New()
   InitializeComponent()
   webBrowser1.Navigate(CType(Me.DataContext, BrowserViewModel).Url)
  End Sub

 End Class
End Namespace
