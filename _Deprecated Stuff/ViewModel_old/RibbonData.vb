Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Namespace ViewModel
 Public Class RibbonData
  Implements INotifyPropertyChanged

  Public ReadOnly Property TabDataCollection() As ObservableCollection(Of TabData)
   Get
    If _tabDataCollection Is Nothing Then
     _tabDataCollection = New ObservableCollection(Of TabData)()
     Dim i As Integer = ViewModelData.TabCount, j As Integer = ViewModelData.ContextualTabGroupCount
     While i > 0
      If j > 0 Then
       Dim contextualTabGroupHeader As String = ContextualTabGroupDataCollection(j - 1).Header
       _tabDataCollection.Insert(0, New TabData("Tab " & i) With {
         .ContextualTabGroupHeader = contextualTabGroupHeader
       })
      Else
       _tabDataCollection.Insert(0, New TabData("Tab " & i))
      End If
      i -= 1
      j -= 1
     End While
    End If
    Return _tabDataCollection
   End Get
  End Property
  Private _tabDataCollection As ObservableCollection(Of TabData)

  Public ReadOnly Property ContextualTabGroupDataCollection() As ObservableCollection(Of ContextualTabGroupData)
   Get
    If _contextualTabGroupDataCollection Is Nothing Then
     _contextualTabGroupDataCollection = New ObservableCollection(Of ContextualTabGroupData)()
     For i As Integer = 0 To ViewModelData.ContextualTabGroupCount - 1
      _contextualTabGroupDataCollection.Add(New ContextualTabGroupData("Grp " & i) With {
        .IsVisible = True
      })
     Next
    End If
    Return _contextualTabGroupDataCollection
   End Get
  End Property
  Private _contextualTabGroupDataCollection As ObservableCollection(Of ContextualTabGroupData)

  Public ReadOnly Property ApplicationMenuData() As MenuButtonData
   Get
    If _applicationMenuData Is Nothing Then
     Dim smallImage As New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative)
     _applicationMenuData = New MenuButtonData(True) With {
      .Label = "AppMenu ",
      .SmallImage = smallImage,
      .ToolTipTitle = "ToolTip Title",
      .ToolTipDescription = "ToolTip Description",
      .ToolTipImage = smallImage
     }
    End If
    Return _applicationMenuData
   End Get
  End Property
  Private _applicationMenuData As MenuButtonData

#Region "INotifyPropertyChanged Members"

  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Protected Sub OnPropertyChanged(e As PropertyChangedEventArgs)
   RaiseEvent PropertyChanged(Me, e)
  End Sub

#End Region

 End Class
End Namespace
