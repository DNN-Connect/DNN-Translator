Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows

Namespace Controls.FolderBrowser
 Public NotInheritable Class TreeViewItemBehaviour

  Private Sub New()
  End Sub

#Region "IsBroughtIntoViewWhenSelected"

  Public Shared Function GetIsBroughtIntoViewWhenSelected(treeViewItem As TreeViewItem) As Boolean
   Return CBool(treeViewItem.GetValue(IsBroughtIntoViewWhenSelectedProperty))
  End Function

  Public Shared Sub SetIsBroughtIntoViewWhenSelected(treeViewItem As TreeViewItem, value As Boolean)
   treeViewItem.SetValue(IsBroughtIntoViewWhenSelectedProperty, value)
  End Sub

  Public Shared ReadOnly IsBroughtIntoViewWhenSelectedProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsBroughtIntoViewWhenSelected", GetType(Boolean), GetType(TreeViewItemBehaviour), New UIPropertyMetadata(False, AddressOf OnIsBroughtIntoViewWhenSelectedChanged))

  Private Shared Sub OnIsBroughtIntoViewWhenSelectedChanged(depObj As DependencyObject, e As DependencyPropertyChangedEventArgs)
   Dim item As TreeViewItem = TryCast(depObj, TreeViewItem)
   If item Is Nothing Then
    Return
   End If

   If TypeOf e.NewValue Is Boolean = False Then
    Return
   End If

   If CBool(e.NewValue) Then
    AddHandler item.Selected, AddressOf item_Selected
   Else
    RemoveHandler item.Selected, AddressOf item_Selected
   End If
  End Sub

  Private Shared Sub item_Selected(sender As Object, e As RoutedEventArgs)
   Dim item As TreeViewItem = TryCast(e.OriginalSource, TreeViewItem)
   If item IsNot Nothing Then
    item.BringIntoView()
   End If
  End Sub

#End Region

 End Class
End Namespace