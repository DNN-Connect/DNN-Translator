Imports DotNetNuke.Translator.ViewModel

Public Class HtmlKeyEdit
 Inherits Window

 Private Sub HtmlKeyView_DataContextChanged(sender As Object, e As System.Windows.DependencyPropertyChangedEventArgs) Handles Me.DataContextChanged
  'teOriginalText.ContentHtml = System.Web.HttpUtility.HtmlDecode(CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.OriginalValue)
  'teTargetText.ContentHtml = System.Web.HttpUtility.HtmlDecode(CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue)
  teOriginalText.ContentHtml = CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.OriginalValue
  teTargetText.ContentHtml = CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue
  CType(Me.DataContext, HtmlKeyViewModel).HasChanges = False
 End Sub

 'Private Sub teTargetText_TextChanged(sender As Object, e As System.EventArgs) Handles teTargetText.SourceUpdated
 ' CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue = System.Web.HttpUtility.HtmlEncode(teTargetText.ContentHtml)
 'End Sub

 Private Sub cmdOk_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdOk.Click
  'CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue = System.Web.HttpUtility.HtmlEncode(teTargetText.ContentHtml)
  CType(Me.DataContext, HtmlKeyViewModel).ResourceKey.TargetValue = teTargetText.ContentHtml
  CType(Me.DataContext, HtmlKeyViewModel).SaveChanges()
  Me.Close()
 End Sub

 Private Sub cmdCancel_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
  Me.Close()
 End Sub

End Class
