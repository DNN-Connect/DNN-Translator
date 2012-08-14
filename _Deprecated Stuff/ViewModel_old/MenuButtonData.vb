Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.Windows.Input
Imports System.ComponentModel

Namespace ViewModel
 Public Class MenuButtonData
  Inherits ControlData
  Public Sub New()
   Me.New(False)
  End Sub

  Public Sub New(isApplicationMenu As Boolean)
   _isApplicationMenu = isApplicationMenu
  End Sub

  Public Property IsVerticallyResizable() As Boolean
   Get
    Return _isVerticallyResizable
   End Get

   Set(value As Boolean)
    If _isVerticallyResizable <> value Then
     _isVerticallyResizable = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsVerticallyResizable"))
    End If
   End Set
  End Property

  Public Property IsHorizontallyResizable() As Boolean
   Get
    Return _isHorizontallyResizable
   End Get

   Set(value As Boolean)
    If _isHorizontallyResizable <> value Then
     _isHorizontallyResizable = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsHorizontallyResizable"))
    End If
   End Set
  End Property

  Private _isVerticallyResizable As Boolean, _isHorizontallyResizable As Boolean

  <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
  Public ReadOnly Property ControlDataCollection() As ObservableCollection(Of ControlData)
   Get
    If _controlDataCollection Is Nothing Then
     _controlDataCollection = New ObservableCollection(Of ControlData)()

     If _nestingDepth <= ViewModelData.MenuItemNestingCount Then
      _nestingDepth += 1

      Dim smallImage As New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative)
      Dim largeImage As New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative)

      For i As Integer = 0 To ViewModelData.GalleryCount - 1
       Dim galleryData As New GalleryData() With {
        .Label = "Gallery " & i,
        .SmallImage = smallImage,
        .LargeImage = largeImage,
        .ToolTipTitle = "ToolTip Title",
        .ToolTipDescription = "ToolTip Description",
        .ToolTipImage = smallImage,
        .Command = ViewModelData.DefaultCommand
       }
       galleryData.SelectedItem = galleryData.CategoryDataCollection(0).GalleryItemDataCollection(ViewModelData.GalleryItemCount - 1)

       _controlDataCollection.Add(galleryData)
      Next
      For i As Integer = 0 To ViewModelData.MenuItemCount - 1
       Dim menuItemData As MenuItemData = If(_isApplicationMenu, New ApplicationMenuItemData(True), New MenuItemData(False))
       menuItemData.Label = "MenuItem " & i
       menuItemData.SmallImage = smallImage
       menuItemData.LargeImage = largeImage
       menuItemData.ToolTipTitle = "ToolTip Title"
       menuItemData.ToolTipDescription = "ToolTip Description"
       menuItemData.ToolTipImage = smallImage
       menuItemData.Command = ViewModelData.DefaultCommand
       menuItemData._nestingDepth = Me._nestingDepth

       _controlDataCollection.Add(menuItemData)
      Next

      _controlDataCollection.Add(New SeparatorData())

      For i As Integer = 0 To ViewModelData.SplitMenuItemCount - 1
       Dim splitMenuItemData As SplitMenuItemData = If(_isApplicationMenu, New ApplicationSplitMenuItemData(True), New SplitMenuItemData(False))
       splitMenuItemData.Label = "SplitMenuItem " & i
       splitMenuItemData.SmallImage = smallImage
       splitMenuItemData.LargeImage = largeImage
       splitMenuItemData.ToolTipTitle = "ToolTip Title"
       splitMenuItemData.ToolTipDescription = "ToolTip Description"
       splitMenuItemData.ToolTipImage = smallImage
       splitMenuItemData.Command = ViewModelData.DefaultCommand
       splitMenuItemData._nestingDepth = Me._nestingDepth
       splitMenuItemData.IsCheckable = True

       _controlDataCollection.Add(splitMenuItemData)
      Next
     End If
    End If
    Return _controlDataCollection
   End Get
  End Property
  Private _controlDataCollection As ObservableCollection(Of ControlData)
  Private _nestingDepth As Integer
  Private _isApplicationMenu As Boolean
 End Class
End Namespace
