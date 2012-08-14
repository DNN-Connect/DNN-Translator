Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows

Namespace Common
 Public Class DataGridContextHelper
  Shared Sub New()
   Dim dp As DependencyProperty = FrameworkElement.DataContextProperty.AddOwner(GetType(DataGridColumn))
   FrameworkElement.DataContextProperty.OverrideMetadata(GetType(DataGrid), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.[Inherits], New PropertyChangedCallback(AddressOf OnDataContextChanged)))
  End Sub

  Public Shared Sub OnDataContextChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
   Dim grid As DataGrid = TryCast(d, DataGrid)
   If grid IsNot Nothing Then
    For Each col As DataGridColumn In grid.Columns
     col.SetValue(FrameworkElement.DataContextProperty, e.NewValue)
    Next
   End If
  End Sub
 End Class
End Namespace