Imports DotNetNuke.Translator.ViewModel

Namespace View
 Public Class HtmlKeyView
  Inherits System.Windows.Controls.UserControl

  Private Sub HtmlKeyView_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
   teOriginalText.Text = System.Web.HttpUtility.HtmlDecode(CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.OriginalValue)
   teTargetText.Text = System.Web.HttpUtility.HtmlDecode(CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue)
  End Sub

  Private Sub teTargetText_TextChanged(sender As Object, e As System.EventArgs) Handles teTargetText.TextChanged
   CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue = System.Web.HttpUtility.HtmlEncode(teTargetText.Text)
  End Sub

 End Class
End Namespace
