Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Namespace ViewModel
 Public Class TabData
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

  Public Property ContextualTabGroupHeader() As String
   Get
    Return _contextualTabGroupHeader
   End Get

   Set(value As String)
    If _contextualTabGroupHeader <> value Then
     _contextualTabGroupHeader = value
     OnPropertyChanged(New PropertyChangedEventArgs("ContextualTabGroupHeader"))
    End If
   End Set
  End Property
  Private _contextualTabGroupHeader As String

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

  Public ReadOnly Property GroupDataCollection() As ObservableCollection(Of GroupData)
   Get
    If _groupDataCollection Is Nothing Then
     Dim smallImage As New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative)
     Dim largeImage As New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative)

     _groupDataCollection = New ObservableCollection(Of GroupData)()
     For i As Integer = 0 To ViewModelData.GroupCount - 1
      _groupDataCollection.Add(New GroupData("Group " & i) With {
       .LargeImage = largeImage,
       .SmallImage = smallImage
      })
     Next
    End If
    Return _groupDataCollection
   End Get
  End Property
  Private _groupDataCollection As ObservableCollection(Of GroupData)

#Region "INotifyPropertyChanged Members"

  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Private Sub OnPropertyChanged(e As PropertyChangedEventArgs)
   RaiseEvent PropertyChanged(Me, e)
  End Sub

#End Region

 End Class
End Namespace
