Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Namespace ViewModel
 Public Class ContextualTabGroupData
  Implements INotifyPropertyChanged

  Public Sub New()
   Me.New(Nothing)
  End Sub

  Public Sub New(header__1 As String)
   Header = header__1
  End Sub


  Public Property Header() As String
   Get
    Return _header
   End Get

   Set(value As String)
    If _header <> value Then
     _header = value
     OnPropertyChanged(New PropertyChangedEventArgs("Header"))
    End If
   End Set
  End Property
  Private _header As String

  Public Property IsVisible() As Boolean
   Get
    Return _isVisible
   End Get

   Set(value As Boolean)
    If _isVisible <> value Then
     _isVisible = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsVisible"))
    End If
   End Set
  End Property
  Private _isVisible As Boolean

  Public ReadOnly Property TabDataCollection() As ObservableCollection(Of TabData)
   Get
    If _tabDataCollection Is Nothing Then
     _tabDataCollection = New ObservableCollection(Of TabData)()
    End If
    Return _tabDataCollection
   End Get
  End Property
  Private _tabDataCollection As ObservableCollection(Of TabData)

#Region "INotifyPropertyChanged Members"

  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Private Sub OnPropertyChanged(e As PropertyChangedEventArgs)
   RaiseEvent PropertyChanged(Me, e)
  End Sub

#End Region

 End Class
End Namespace
