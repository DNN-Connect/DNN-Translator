Namespace Common
 Public Class BindingProxy
  Inherits Freezable
#Region "Overrides of Freezable"

  Protected Overrides Function CreateInstanceCore() As Freezable
   Return New BindingProxy()
  End Function

#End Region

  Public Property Data() As Object
   Get
    Return DirectCast(GetValue(DataProperty), Object)
   End Get
   Set(value As Object)
    SetValue(DataProperty, value)
   End Set
  End Property

  ' Using a DependencyProperty as the backing store for Data.  This enables animation, styling, binding, etc...
  Public Shared ReadOnly DataProperty As DependencyProperty = DependencyProperty.Register("Data", GetType(Object), GetType(BindingProxy), New UIPropertyMetadata(Nothing))
 End Class
End Namespace