Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel

Namespace ViewModel
 Public Class RadioButtonData
  Inherits ControlData
  Public Property IsChecked() As Boolean
   Get
    Return _isChecked
   End Get

   Set(value As Boolean)
    If _isChecked <> value Then
     _isChecked = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsChecked"))
    End If
   End Set
  End Property
  Private _isChecked As Boolean
 End Class
End Namespace
