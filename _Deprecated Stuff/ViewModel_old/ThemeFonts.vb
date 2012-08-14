Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel

Namespace ViewModel
 ''' <summary>
 ''' DataModel for ThemeFonts Gallery
 ''' </summary>
 Public Class ThemeFonts
  Implements INotifyPropertyChanged

  Public Property Label() As String
   Get
    Return _label
   End Get

   Set(value As String)
    If _label <> value Then
     _label = value
     OnPropertyChanged(New PropertyChangedEventArgs("Label"))
    End If
   End Set
  End Property
  Private _label As String


  Public Property Field1() As String
   Get
    Return _field1
   End Get

   Set(value As String)
    If _field1 <> value Then
     _field1 = value
     OnPropertyChanged(New PropertyChangedEventArgs("Field1"))
    End If
   End Set
  End Property
  Private _field1 As String

  Public Property Field2() As String
   Get
    Return _field2
   End Get

   Set(value As String)
    If _field2 <> value Then
     _field2 = value
     OnPropertyChanged(New PropertyChangedEventArgs("Field2"))
    End If
   End Set
  End Property
  Private _field2 As String

  Public Property Field3() As String
   Get
    Return _field3
   End Get

   Set(value As String)
    If _field3 <> value Then
     _field3 = value
     OnPropertyChanged(New PropertyChangedEventArgs("Field3"))
    End If
   End Set
  End Property
  Private _field3 As String

#Region "INotifyPropertyChanged Members"

  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Protected Sub OnPropertyChanged(e As PropertyChangedEventArgs)
   RaiseEvent PropertyChanged(Me, e)
  End Sub

#End Region

 End Class
End Namespace
