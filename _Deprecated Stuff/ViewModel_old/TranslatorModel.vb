Imports DotNetNuke.Translator.Common

Namespace ViewModel
 Public Class TranslatorModel

#Region " Menu "
  Public Shared ReadOnly Property Open() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Open"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Open (Ctrl+O)"

      Dim menuItemData As New MenuItemData() With {
        .Label = Str,
        .SmallImage = New Uri("Images\SmallIcon.png", UriKind.Relative),
        .ToolTipTitle = TooTipTitle,
      .Command = New DelegateCommand(AddressOf OpenClicked, AddressOf DefaultCanExecute)
      }
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RecentDocuments() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Recent Locations"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim galleryCategoryData As New GalleryCategoryData(Of RecentDocumentData)()
      For Each recentLoc As KeyValuePair(Of Integer, String) In _recentLocations.RecentLocations
       Dim recentDocumentData As New RecentDocumentData() With {
         .Index = recentLoc.Key,
         .Label = recentLoc.Value
       }
       galleryCategoryData.GalleryItemDataCollection.Add(recentDocumentData)
      Next
      _dataCollection(Str) = galleryCategoryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property
#End Region

#Region " Opening "
  Private Shared Sub OpenClicked()

   Dim fbd As New System.Windows.Forms.FolderBrowserDialog
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    If IO.Directory.Exists(fbd.SelectedPath) Then
     RaiseEvent OpenNewLocation(fbd.SelectedPath)
    End If
   End If

  End Sub

  Public Shared Event OpenNewLocation(location As String)
  Public Shared Event LocationChanged()

  Public Shared Sub OpenedNewLocation(location As String)
   _currentLocation = location
   _recentLocations.Push(location)
   Data = New TranslatorData(location)
   RaiseEvent LocationChanged()
  End Sub
#End Region

#Region " Exiting "
  Public Shared Sub ExitApplication()
   _recentLocations.Save()
   Try
    Data.Save()
   Catch ex As Exception
   End Try
  End Sub
#End Region

#Region "Data"

  Private Const HelpFooterTitle As String = "Press F1 for more help."
  Private Shared _lockObject As New Object()
  Private Shared _dataCollection As New Dictionary(Of String, ControlData)()

  ' Store any data that doesnt inherit from ControlData
  Private Shared _miscData As New Dictionary(Of String, Object)()

  Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
   target = value
   Return value
  End Function

  Private Shared _recentLocations As New Common.RecentLocations
  Private Shared _currentLocation As String = ""
  Public Shared Property Data As TranslatorData

#End Region

  Private Shared Sub DefaultExecuted()
  End Sub

  Private Shared Function DefaultCanExecute() As Boolean
   Return True
  End Function

 End Class
End Namespace
