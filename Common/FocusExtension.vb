Namespace Common
 Public Class FocusExtension
  Public Shared Function GetIsFocused(obj As DependencyObject) As Boolean
   Return CBool(obj.GetValue(IsFocusedProperty))
  End Function

  Public Shared Sub SetIsFocused(obj As DependencyObject, value As Boolean)
   obj.SetValue(IsFocusedProperty, value)
  End Sub

  Public Shared ReadOnly IsFocusedProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsFocused", GetType(Boolean), GetType(FocusExtension), New UIPropertyMetadata(False, AddressOf OnIsFocusedPropertyChanged))

  Private Shared Sub OnIsFocusedPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
   Dim uie = DirectCast(d, UIElement)
   If CBool(e.NewValue) Then
    uie.Focus()
   End If
  End Sub
 End Class
End Namespace
