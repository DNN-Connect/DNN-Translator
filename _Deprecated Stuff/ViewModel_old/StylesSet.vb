Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel

Namespace ViewModel
 ''' <summary>
 ''' DataModel for Styles Gallery.
 ''' </summary>
 Public Class StylesSet
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

  Public Property IsSelected() As Boolean
   Get
    Return _isSelected
   End Get

   Set(value As Boolean)
    If _isSelected <> value Then
     _isSelected = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsSelected"))
    End If
   End Set
  End Property
  Private _isSelected As Boolean

#Region "INotifyPropertyChanged Members"

  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Protected Sub OnPropertyChanged(e As PropertyChangedEventArgs)
   RaiseEvent PropertyChanged(Me, e)
  End Sub

#End Region

 End Class
End Namespace
