Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.Windows.Input
Imports System.ComponentModel

Namespace ViewModel
 Public Class GalleryCategoryData
  Inherits ControlData
  <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
  Public ReadOnly Property GalleryItemDataCollection() As ObservableCollection(Of GalleryItemData)
   Get
    If _controlDataCollection Is Nothing Then
     _controlDataCollection = New ObservableCollection(Of GalleryItemData)()

     Dim smallImage As New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative)
     Dim largeImage As New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative)

     For i As Integer = 0 To ViewModelData.GalleryItemCount - 1
      _controlDataCollection.Add(New GalleryItemData() With {
        .Label = "GalleryItem " & i,
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
  Private _controlDataCollection As ObservableCollection(Of GalleryItemData)
 End Class

 Public Class GalleryCategoryData(Of T)
  Inherits ControlData
  <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
  Public ReadOnly Property GalleryItemDataCollection() As ObservableCollection(Of T)
   Get
    If _controlDataCollection Is Nothing Then
     _controlDataCollection = New ObservableCollection(Of T)()
    End If
    Return _controlDataCollection
   End Get
  End Property
  Private _controlDataCollection As ObservableCollection(Of T)
 End Class
End Namespace
