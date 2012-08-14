Imports System.Globalization

Namespace Common
 Public Class BoolToVisibilityConverter
  Implements IValueConverter
  Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements System.Windows.Data.IValueConverter.Convert
   Try
    Dim v = CBool(value)
    Return If(v, Visibility.Visible, Visibility.Collapsed)
   Catch generatedExceptionName As InvalidCastException
    Return Visibility.Collapsed
   End Try
  End Function

  Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements System.Windows.Data.IValueConverter.ConvertBack
   Throw New NotImplementedException()
  End Function
 End Class
End Namespace
