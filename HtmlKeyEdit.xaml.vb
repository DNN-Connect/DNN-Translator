Imports DotNetNuke.Translator.ViewModel

Public Class HtmlKeyEdit
 Inherits Window

 Private Sub HtmlKeyView_DataContextChanged(sender As Object, e As System.Windows.DependencyPropertyChangedEventArgs) Handles Me.DataContextChanged
  teOriginalText.Text = System.Web.HttpUtility.HtmlDecode(CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.OriginalValue)
  teTargetText.Text = System.Web.HttpUtility.HtmlDecode(CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue)
  CType(Me.DataContext, HtmlKeyViewModel).HasChanges = False
 End Sub

 Private Sub teTargetText_TextChanged(sender As Object, e As System.EventArgs) Handles teTargetText.TextChanged
  CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue = System.Web.HttpUtility.HtmlEncode(teTargetText.Text)
 End Sub

 Private Sub cmdOk_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdOk.Click
  CType(Me.DataContext, HtmlKeyViewModel).SaveChanges()
  Me.Close()
 End Sub

 Private Sub cmdCancel_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
  Me.Close()
 End Sub

End Class
