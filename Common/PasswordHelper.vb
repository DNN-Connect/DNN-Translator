Namespace Common
 Public NotInheritable Class PasswordHelper
  Private Sub New()
  End Sub

  Public Shared ReadOnly PasswordProperty As DependencyProperty = DependencyProperty.RegisterAttached("Password", GetType(String), GetType(PasswordHelper), New FrameworkPropertyMetadata(String.Empty, AddressOf OnPasswordPropertyChanged))

  Public Shared ReadOnly AttachProperty As DependencyProperty = DependencyProperty.RegisterAttached("Attach", GetType(Boolean), GetType(PasswordHelper), New PropertyMetadata(False, AddressOf Attach))

  Private Shared ReadOnly IsUpdatingProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsUpdating", GetType(Boolean), GetType(PasswordHelper))

  Public Shared Sub SetAttach(dp As DependencyObject, value As Boolean)
   dp.SetValue(AttachProperty, value)
  End Sub

  Public Shared Function GetAttach(dp As DependencyObject) As Boolean
   Return CBool(dp.GetValue(AttachProperty))
  End Function

  Public Shared Function GetPassword(dp As DependencyObject) As String
   Return DirectCast(dp.GetValue(PasswordProperty), String)
  End Function

  Public Shared Sub SetPassword(dp As DependencyObject, value As String)
   dp.SetValue(PasswordProperty, value)
  End Sub

  Private Shared Function GetIsUpdating(dp As DependencyObject) As Boolean
   Return CBool(dp.GetValue(IsUpdatingProperty))
  End Function

  Private Shared Sub SetIsUpdating(dp As DependencyObject, value As Boolean)
   dp.SetValue(IsUpdatingProperty, value)
  End Sub

  Private Shared Sub OnPasswordPropertyChanged(sender As DependencyObject, e As DependencyPropertyChangedEventArgs)
   Dim passwordBox As PasswordBox = TryCast(sender, PasswordBox)
   RemoveHandler passwordBox.PasswordChanged, AddressOf PasswordChanged

   If Not CBool(GetIsUpdating(passwordBox)) Then
    passwordBox.Password = DirectCast(e.NewValue, String)
   End If
   AddHandler passwordBox.PasswordChanged, AddressOf PasswordChanged
  End Sub

  Private Shared Sub Attach(sender As DependencyObject, e As DependencyPropertyChangedEventArgs)
   Dim passwordBox As PasswordBox = TryCast(sender, PasswordBox)

   If passwordBox Is Nothing Then
    Return
   End If

   If CBool(e.OldValue) Then
    RemoveHandler passwordBox.PasswordChanged, AddressOf PasswordChanged
   End If

   If CBool(e.NewValue) Then
    AddHandler passwordBox.PasswordChanged, AddressOf PasswordChanged
   End If
  End Sub

  Private Shared Sub PasswordChanged(sender As Object, e As RoutedEventArgs)
   Dim passwordBox As PasswordBox = TryCast(sender, PasswordBox)
   SetIsUpdating(passwordBox, True)
   SetPassword(passwordBox, passwordBox.Password)
   SetIsUpdating(passwordBox, False)
  End Sub
 End Class
End Namespace