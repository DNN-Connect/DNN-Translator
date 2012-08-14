Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Input

Namespace Controls.FolderBrowser
 Public NotInheritable Class InputBindingsManager
  Private Sub New()
  End Sub

  Public Shared ReadOnly UpdatePropertySourceWhenEnterPressedProperty As DependencyProperty = DependencyProperty.RegisterAttached("UpdatePropertySourceWhenEnterPressed", GetType(DependencyProperty), GetType(InputBindingsManager), New PropertyMetadata(Nothing, AddressOf OnUpdatePropertySourceWhenEnterPressedPropertyChanged))

  Public Shared Sub SetUpdatePropertySourceWhenEnterPressed(dp As DependencyObject, value As DependencyProperty)
   dp.SetValue(UpdatePropertySourceWhenEnterPressedProperty, value)
  End Sub

  Public Shared Function GetUpdatePropertySourceWhenEnterPressed(dp As DependencyObject) As DependencyProperty
   Return DirectCast(dp.GetValue(UpdatePropertySourceWhenEnterPressedProperty), DependencyProperty)
  End Function

  Private Shared Sub OnUpdatePropertySourceWhenEnterPressedPropertyChanged(dp As DependencyObject, e As DependencyPropertyChangedEventArgs)
   Dim element As UIElement = TryCast(dp, UIElement)
   If element Is Nothing Then
    Return
   End If
   If e.OldValue IsNot Nothing Then
    RemoveHandler element.PreviewKeyDown, AddressOf HandlePreviewKeyDown
   End If
   If e.NewValue IsNot Nothing Then
    AddHandler element.PreviewKeyDown, New KeyEventHandler(AddressOf HandlePreviewKeyDown)
   End If
  End Sub

  Private Shared Sub HandlePreviewKeyDown(sender As Object, e As KeyEventArgs)
   If e.Key = Key.Enter Then
    DoUpdateSource(e.Source)
   End If
  End Sub

  Private Shared Sub DoUpdateSource(source As Object)
   Dim [property] As DependencyProperty = GetUpdatePropertySourceWhenEnterPressed(TryCast(source, DependencyObject))
   If [property] Is Nothing Then
    Return
   End If
   Dim elt As UIElement = TryCast(source, UIElement)
   If elt Is Nothing Then
    Return
   End If
   Dim binding As BindingExpression = BindingOperations.GetBindingExpression(elt, [property])
   If binding IsNot Nothing Then
    binding.UpdateSource()
   End If
  End Sub
 End Class
End Namespace
