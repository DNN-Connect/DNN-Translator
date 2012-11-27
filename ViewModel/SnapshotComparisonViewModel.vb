Namespace ViewModel
 Public Class SnapshotComparisonViewModel
  Inherits WorkspaceViewModel

  Public Property MainWindow As MainWindowViewModel = Nothing
  Public Property Start As String = ""
  Public Property ComparisonFile As String = ""
  Public Property ComparisonSnapshot As Snapshot = Nothing
  Public Property CurrentSnapshot As Snapshot = Nothing
  Public Property NewKeys As New Common.SortableObservableCollection(Of ResourceKeyViewModel)
  Public Property DeprecatedKeys As New Common.SortableObservableCollection(Of ResourceKeyViewModel)
  Public Property ChangedKeys As New Common.SortableObservableCollection(Of ResourceKeyViewModel)
  Public Property NewWords As Integer = 0
  Public Property Locale As String = ""

  Public Sub New(mainWindow As MainWindowViewModel, start As String, comparisonFile As String, locale As String)
   Me.MainWindow = mainWindow
   Me.Start = start
   Me.ComparisonFile = comparisonFile
   Me.Locale = locale
   Me.ID = String.Format("{0}_{1}_{2}", start, comparisonFile, locale)
   Me.ComparisonSnapshot = New Snapshot
   Me.ComparisonSnapshot.ReadXml(comparisonFile, System.Data.XmlReadMode.IgnoreSchema)
   Me.CurrentSnapshot = mainWindow.ProjectSettings.CurrentSnapShot

   Dim allRows As New List(Of String)
   For Each resFile As Common.ResourceFile In From f In CurrentSnapshot.ResourceFiles.Values Where f.FilePath.StartsWith(start) Select f
    For Each resKey As Common.Resource In resFile.Resources.Values
     Dim fileKey As String = resKey.FileKey
     Dim key As String = resKey.Key
     Dim origRow As IEnumerable(Of Snapshot.ResourcesRow) = From x In ComparisonSnapshot.Resources Where x.FileKey = fileKey And x.ResourceKey = key Select x
     If origRow.Count > 0 Then
      'If System.Web.HttpUtility.HtmlDecode(resKey.Value) <> origRow(0).ResourceValue Then
      If resKey.Value <> origRow(0).ResourceValue Then
       ChangedKeys.Add(New ResourceKeyViewModel(resKey, Common.Globals.GetTranslation(mainWindow.ProjectSettings.Location, resKey.FileKey, resKey.Key, locale), origRow(0).ResourceValue))
       NewWords += resKey.Value.Length - resKey.Value.Replace(" ", "").Length
      End If
     Else
      NewKeys.Add(New ResourceKeyViewModel(resKey, Common.Globals.GetTranslation(mainWindow.ProjectSettings.Location, resKey.FileKey, resKey.Key, locale)))
      NewWords += resKey.Value.Length - resKey.Value.Replace(" ", "").Length
     End If
     allRows.Add(String.Format("{0}/{1}", resKey.FileKey, resKey.Key))
    Next
   Next
   For Each crtRow As Snapshot.ResourcesRow In CurrentSnapshot.Resources.Rows
    Dim fileKey As String = crtRow.FileKey
    Dim resourceKey As String = crtRow.ResourceKey
   Next
   For Each cmpRow As Snapshot.ResourcesRow In ComparisonSnapshot.Resources.Rows
    If Not allRows.Contains(String.Format("{0}/{1}", cmpRow.FileKey, cmpRow.ResourceKey)) Then
     Dim mockOrig As New Common.Resource With {.FileKey = cmpRow.FileKey, .Key = cmpRow.ResourceKey, .Value = ""}
     DeprecatedKeys.Add(New ResourceKeyViewModel(mockOrig, Common.Globals.GetTranslation(mainWindow.ProjectSettings.Location, mockOrig.FileKey, mockOrig.Key, locale), cmpRow.ResourceValue))
    End If
   Next
  End Sub

  Private _openResourcesCommand As RelayCommand
  Public ReadOnly Property OpenResourcesCommand As RelayCommand
   Get
    If _openResourcesCommand Is Nothing Then
     _openResourcesCommand = New RelayCommand(Sub(param) Me.OpenResourcesClicked(param))
    End If
    Return _openResourcesCommand
   End Get
  End Property

  Protected Sub OpenResourcesClicked(param As Object)
   Dim what As String = CStr(param).ToLower
   Select Case what
    Case "new"
     Dim newID As String = String.Format("Newkeys_{0}_{1}", Me.ID, MainWindow.SelectedLocale.Name)
     For Each ws As WorkspaceViewModel In MainWindow.Workspaces
      If ws.ID = newID Then
       MainWindow.SetActiveWorkspace(ws)
       Exit Sub
      End If
     Next
     Dim resEdit As New ResourceCollectionViewModel(MainWindow)
     resEdit.ShowFileColumn = True
     resEdit.AddKeys(NewKeys)
     resEdit.HasChanges = False
     resEdit.ID = newID
     resEdit.DisplayName = "New Texts"
     resEdit.CompareColumnHeader = ComparisonFile.Substring(ComparisonFile.LastIndexOf("\"c) + 1)
     MainWindow.Workspaces.Add(resEdit)
     MainWindow.SetActiveWorkspace(resEdit)
    Case "changed"
     Dim newID As String = String.Format("Changedkeys_{0}_{1}", Me.ID, MainWindow.SelectedLocale.Name)
     For Each ws As WorkspaceViewModel In MainWindow.Workspaces
      If ws.ID = newID Then
       MainWindow.SetActiveWorkspace(ws)
       Exit Sub
      End If
     Next
     Dim resEdit As New ResourceCollectionViewModel(MainWindow)
     resEdit.ShowFileColumn = True
     resEdit.AddKeys(ChangedKeys)
     resEdit.HasChanges = False
     resEdit.ID = newID
     resEdit.DisplayName = "Changed Texts"
     resEdit.CompareColumnHeader = ComparisonFile.Substring(ComparisonFile.LastIndexOf("\"c) + 1)
     MainWindow.Workspaces.Add(resEdit)
     MainWindow.SetActiveWorkspace(resEdit)
   End Select

  End Sub

 End Class
End Namespace