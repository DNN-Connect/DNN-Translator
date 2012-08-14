Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Input

Namespace ViewModel
 Public Class GalleryData
  Inherits ControlData
  <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
  Public ReadOnly Property CategoryDataCollection() As ObservableCollection(Of GalleryCategoryData)
   Get
    If _controlDataCollection Is Nothing Then
     _controlDataCollection = New ObservableCollection(Of GalleryCategoryData)()

     Dim smallImage As New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative)
     Dim largeImage As New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative)

     For i As Integer = 0 To ViewModelData.GalleryCategoryCount - 1
      _controlDataCollection.Add(New GalleryCategoryData() With {
       .Label = "GalleryCategory " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
    End If
    Return _controlDataCollection
   End Get
  End Property
  Private _controlDataCollection As ObservableCollection(Of GalleryCategoryData)

  Public Property SelectedItem() As GalleryItemData
   Get
    Return _selectedItem
   End Get
   Set(value As GalleryItemData)
    If _selectedItem IsNot value Then
     _selectedItem = value
     OnPropertyChanged(New PropertyChangedEventArgs("SelectedItem"))
    End If
   End Set
  End Property
  Private _selectedItem As GalleryItemData

  Public Property CanUserFilter() As Boolean
   Get
    Return _canUserFilter
   End Get

   Set(value As Boolean)
    If _canUserFilter <> value Then
     _canUserFilter = value
     OnPropertyChanged(New PropertyChangedEventArgs("CanUserFilter"))
    End If
   End Set
  End Property

  Private _canUserFilter As Boolean
 End Class

 Public Class GalleryData(Of T)
  Inherits ControlData
  <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
  Public ReadOnly Property CategoryDataCollection() As ObservableCollection(Of GalleryCategoryData(Of T))
   Get
    If _controlDataCollection Is Nothing Then
     _controlDataCollection = New ObservableCollection(Of GalleryCategoryData(Of T))()
    End If
    Return _controlDataCollection
   End Get
  End Property
  Private _controlDataCollection As ObservableCollection(Of GalleryCategoryData(Of T))

  Public Property SelectedItem() As T
   Get
    Return _selectedItem
   End Get
   Set(value As T)
    If Not [Object].Equals(value, _selectedItem) Then
     _selectedItem = value
     OnPropertyChanged(New PropertyChangedEventArgs("SelectedItem"))
    End If
   End Set
  End Property
  Private _selectedItem As T

  Public Property CanUserFilter() As Boolean
   Get
    Return _canUserFilter
   End Get

   Set(value As Boolean)
    If _canUserFilter <> value Then
     _canUserFilter = value
     OnPropertyChanged(New PropertyChangedEventArgs("CanUserFilter"))
    End If
   End Set
  End Property

  Private _canUserFilter As Boolean
 End Class
End Namespace
