Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel

Namespace ViewModel
 Public Class TextBoxData
  Inherits ControlData
  Public Property Text() As String
   Get
    Return _text
   End Get

   Set(value As String)
    If _text <> value Then
     _text = value
     OnPropertyChanged(New PropertyChangedEventArgs("Text"))
    End If
   End Set
  End Property
  Private _text As String
 End Class
End Namespace
