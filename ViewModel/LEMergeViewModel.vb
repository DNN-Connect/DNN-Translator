Imports DotNetNuke.Translator.Common.LEService
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Namespace ViewModel
 Public Class LEMergeViewModel
  Inherits WorkspaceViewModel

#Region " Properties "
  Public Property MainWindow As MainWindowViewModel
  Public Property RemoteResources As List(Of TextInfo)
  Public Property FileList As New List(Of String)
  Public Property LocalData As New Dictionary(Of String, Common.ResourceFile)
  Public Property service As Common.LEService.LEService
  Public Property ObjectName As String
  Public Property ObjectVersion As String
  Public Property Start As String
  Public Property SameKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property MissingLocalKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property MissingRemoteKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property DifferentKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property TotalLocalResources As Integer = 0
  Private _backgroundWorker As BackgroundWorker
  Public Property FailedCompare As Boolean = False
#End Region

#Region " Constructor "
  Public Sub New(parameters As Common.ParameterList)
   Me.MainWindow = CType(parameters.ParentWindow, MainWindowViewModel)
   Me.ObjectName = parameters.Params(0)
   Me.ObjectVersion = parameters.Params(1)
   Me.Start = parameters.Params(2)
   Me.DisplayName = String.Format("Merge {0}", ObjectName)
   Me.ID = Me.DisplayName

   ' download current situation from server
   MainWindow.BusyMessage = "Downloading from server"
   service = New Common.LEService.LEService(MainWindow.ProjectSettings.ConnectionUrl, MainWindow.ProjectSettings.Username, MainWindow.ProjectSettings.Password)
   Try
    RemoteResources = service.GetResources(ObjectName, ObjectVersion, MainWindow.ProjectSettings.MappedLocale, Start)
   Catch ex As Exception
    If service.IsInError Then
     MsgBox(service.LastErrorMessage, MsgBoxStyle.OkOnly)
    Else
     MsgBox(ex.Message, MsgBoxStyle.Critical)
    End If
    FailedCompare = True
    Exit Sub
   End Try
   If RemoteResources Is Nothing OrElse RemoteResources.Count = 0 Then
    MsgBox("No resources found", MsgBoxStyle.OkOnly)
    RemoteResources = New List(Of TextInfo)
    FailedCompare = True
    Exit Sub
   End If

   ' load res files
   MainWindow.BusyMessage = "Loading local data"
   ' get list of files
   For Each ti As TextInfo In RemoteResources
    If Not FileList.Contains(ti.FilePath) Then FileList.Add(ti.FilePath)
   Next

   Compare()

  End Sub
#End Region

#Region " Comparing "
  Private Sub Compare()

   LocalData = New Dictionary(Of String, Common.ResourceFile)
   TotalLocalResources = 0
   SameKeys = New ObservableCollection(Of ResourceKeyViewModel)
   MissingLocalKeys = New ObservableCollection(Of ResourceKeyViewModel)
   MissingRemoteKeys = New ObservableCollection(Of ResourceKeyViewModel)
   DifferentKeys = New ObservableCollection(Of ResourceKeyViewModel)

   ' load local data
   For Each file As String In FileList
    MainWindow.BusyMessage = String.Format("Loading {0}", file)
    LocalData.Add(file, New Common.ResourceFile(file, Common.Globals.GetResourceFilePath(MainWindow.ProjectSettings.Location, file, MainWindow.ProjectSettings.TargetLocale)))
    TotalLocalResources += LocalData(file).Resources.Count
   Next

   Try
    ' compare with what we have
    MainWindow.BusyMessage = "Comparing with local data"
    For Each ti As TextInfo In RemoteResources
     MainWindow.BusyMessage = String.Format("Comparing {0}:{1}", ti.FilePath, ti.TextKey)
     Dim localValue As Common.Resource = Nothing
     If LocalData(ti.FilePath) IsNot Nothing AndAlso LocalData(ti.FilePath).Resources.ContainsKey(ti.TextKey) Then localValue = LocalData(ti.FilePath).Resources(ti.TextKey)
     If localValue IsNot Nothing AndAlso localValue.Value = ti.Translation Then
      SameKeys.Add(New ResourceKeyViewModel(localValue, ti))
     ElseIf localValue Is Nothing Then ' empty local value
      MissingLocalKeys.Add(New ResourceKeyViewModel(localValue, ti))
     ElseIf ti.Translation = "" Then ' empty remote value
      MissingRemoteKeys.Add(New ResourceKeyViewModel(localValue, ti))
     Else ' different local value
      Dim dk As New ResourceKeyViewModel(localValue, ti)
      If ti.LastModified > LocalData(ti.FilePath).Resources(ti.TextKey).LastModified Then
       dk.HighlightCompareValue = True
      Else
       dk.HighlightTargetValue = True
      End If
      DifferentKeys.Add(dk)
     End If
    Next
   Catch ex As Exception

   End Try

   Me.OnPropertyChanged("SameKeys")
   Me.OnPropertyChanged("MissingLocalKeys")
   Me.OnPropertyChanged("MissingRemoteKeys")
   Me.OnPropertyChanged("DifferentKeys")
   Me.OnPropertyChanged("LocalData")
   Me.OnPropertyChanged("TotalLocalResources")

  End Sub
#End Region

#Region " Opening Resources "
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

   Dim newID As String = String.Format(CType(param, String).ToLower & "keys_{0}_{1}", Me.ID, MainWindow.SelectedLocale.Name)
   For Each ws As WorkspaceViewModel In MainWindow.Workspaces
    If ws.ID = newID Then
     MainWindow.SetActiveWorkspace(ws)
     Exit Sub
    End If
   Next

   MainWindow.IsBusy = True
   MainWindow.BusyMessage = "Opening Resources"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf OpenResources
   AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf OpenResourcesCompleted
   Dim p As New Common.ParameterList(Me.MainWindow)
   p.Params.Add(CType(param, String))
   _backgroundWorker.RunWorkerAsync(p)

  End Sub

  Private Sub OpenResources(sender As Object, e As DoWorkEventArgs)
   Dim p As Common.ParameterList = CType(e.Argument, Common.ParameterList)
   Dim what As String = p.Params(0).ToLower
   Select Case what
    Case "same"
     e.Result = ShowResourceKeys(what, "Same Texts", SameKeys)
    Case "missinglocal"
     e.Result = ShowResourceKeys(what, "Missing Local Texts", MissingLocalKeys)
    Case "missingremote"
     e.Result = ShowResourceKeys(what, "Missing Remote Texts", MissingRemoteKeys)
    Case "different"
     e.Result = ShowResourceKeys(what, "Different Texts", DifferentKeys)
   End Select
  End Sub

  Private Sub OpenResourcesCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   Dim resEdit As ResourceCollectionViewModel = CType(e.Result, ResourceCollectionViewModel)
   MainWindow.Workspaces.Add(resEdit)
   MainWindow.SetActiveWorkspace(resEdit)
   MainWindow.IsBusy = False
  End Sub

  Private Function ShowResourceKeys(type As String, displayName As String, keyCollection As ObservableCollection(Of ResourceKeyViewModel)) As Object
   Dim newID As String = String.Format(type & "keys_{0}_{1}", Me.ID, MainWindow.SelectedLocale.Name)
   Dim resEdit As New ResourceCollectionViewModel(MainWindow)
   resEdit.ShowFileColumn = True
   resEdit.AddKeys(keyCollection)
   resEdit.HasChanges = False
   resEdit.ID = newID
   resEdit.DisplayName = displayName
   resEdit.IsRemoteResourceCollection = True
   resEdit.CompareColumnHeader = "Remote Values"
   Return resEdit
  End Function
#End Region

#Region " Uploading "
  Private _UploadAllCommand As RelayCommand
  Public ReadOnly Property UploadAllCommand As RelayCommand
   Get
    If _UploadAllCommand Is Nothing Then
     _UploadAllCommand = New RelayCommand(Sub(param) Me.UploadAllClicked(param))
    End If
    Return _UploadAllCommand
   End Get
  End Property

  Protected Sub UploadAllClicked(param As Object)
   MainWindow.IsBusy = True
   MainWindow.BusyMessage = "Uploading Resources"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf UploadAll
   AddHandler _backgroundWorker.RunWorkerCompleted, Sub(sender As Object, e As RunWorkerCompletedEventArgs) MainWindow.IsBusy = False
   _backgroundWorker.RunWorkerAsync(param)
  End Sub

  Private Sub UploadAll(sender As Object, e As DoWorkEventArgs)

   Dim toupload As New List(Of TextInfo)
   For Each rkv As ResourceKeyViewModel In MissingRemoteKeys
    toupload.Add(rkv.LEText)
   Next
   For Each rkv As ResourceKeyViewModel In DifferentKeys
    toupload.Add(rkv.LEText)
   Next
   Dim service As New Common.LEService.LEService(MainWindow.ProjectSettings.ConnectionUrl, MainWindow.ProjectSettings.Username, MainWindow.ProjectSettings.Password)
   Try
    service.UploadResources(toupload)
   Catch ex As Exception
    If service.IsInError Then
     MsgBox(service.LastErrorMessage, MsgBoxStyle.OkOnly)
    Else
     MsgBox(ex.Message, MsgBoxStyle.Critical)
    End If
   End Try

   ' re-compare with what we have
   Compare()

  End Sub
#End Region

#Region " Downloading "
  Private _DownloadAllCommand As RelayCommand
  Public ReadOnly Property DownloadAllCommand As RelayCommand
   Get
    If _DownloadAllCommand Is Nothing Then
     _DownloadAllCommand = New RelayCommand(Sub(param) Me.DownloadAllClicked(param))
    End If
    Return _DownloadAllCommand
   End Get
  End Property

  Protected Sub DownloadAllClicked(param As Object)
   MainWindow.IsBusy = True
   MainWindow.BusyMessage = "Downloading Resources"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf DownloadAll
   AddHandler _backgroundWorker.RunWorkerCompleted, Sub(sender As Object, e As RunWorkerCompletedEventArgs) MainWindow.IsBusy = False
   _backgroundWorker.RunWorkerAsync(param)
  End Sub

  Private Sub DownloadAll(sender As Object, e As DoWorkEventArgs)
   ' get a list of files
   Dim fileList As New List(Of String)
   For Each k As ResourceKeyViewModel In MissingLocalKeys
    If Not fileList.Contains(k.ResourceFile) Then
     fileList.Add(k.ResourceFile)
    End If
   Next
   For Each k As ResourceKeyViewModel In DifferentKeys
    If Not fileList.Contains(k.ResourceFile) Then
     fileList.Add(k.ResourceFile)
    End If
   Next
   ' handle each file
   For Each resFile As String In fileList
    Dim fileKey As String = resFile
    Dim targetFile As String = MainWindow.ProjectSettings.Location & "\" & resFile.Replace(".resx", "." & MainWindow.ProjectSettings.TargetLocale & ".resx")
    Dim targetResx As New Common.ResourceFile(fileKey, targetFile)
    For Each k As ResourceKeyViewModel In From x In MissingLocalKeys Where x.ResourceFile = fileKey Select x
     targetResx.SetResourceValue(k.Key, k.LEText.Translation, k.LEText.LastModified)
     k.TargetValue = k.LEText.Translation
    Next
    For Each k As ResourceKeyViewModel In From x In DifferentKeys Where x.ResourceFile = fileKey Select x
     targetResx.SetResourceValue(k.Key, k.LEText.Translation, k.LEText.LastModified)
     k.TargetValue = k.LEText.Translation
    Next
    targetResx.Save()
   Next

   ' re-compare with what we have
   Compare()

  End Sub
#End Region

 End Class
End Namespace
