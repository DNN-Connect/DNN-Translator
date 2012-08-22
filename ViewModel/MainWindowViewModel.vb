Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Imports DotNetNuke.Translator.Common

Namespace ViewModel
 Public Class MainWindowViewModel
  Inherits WorkspaceViewModel

#Region " Other Properties "
  Private _currentLocation As String = ""
#End Region

#Region " Constructor "
  Public Sub New()
   SetLocale(TranslatorSettings.DefaultTargetLocale)
  End Sub
#End Region

#Region " Busy Overlay System "
  Private _backgroundWorker As BackgroundWorker
  Private _isBusy As Boolean
  Public Property IsBusy() As Boolean
   Get
    Return _isBusy
   End Get
   Set(ByVal value As Boolean)
    _isBusy = value
    Me.OnPropertyChanged("IsBusy")
   End Set
  End Property

  Private _BusyMessage As String
  Public Property BusyMessage() As String
   Get
    Return _BusyMessage
   End Get
   Set(ByVal value As String)
    _BusyMessage = value
    Me.OnPropertyChanged("BusyMessage")
   End Set
  End Property
#End Region

#Region " Keyboard Handling "
  Private _keyboardCommand As RelayCommand
  Public ReadOnly Property KeyboardCommand As RelayCommand
   Get
    If _keyboardCommand Is Nothing Then
     _keyboardCommand = New RelayCommand(Sub(param) KeyboardCommandHandler(param))
    End If
    Return _keyboardCommand
   End Get
  End Property

  Protected Sub KeyboardCommandHandler(param As Object)
   Dim command As String = CType(param, String)
   Select Case command
    Case "SelectAll"
     Dim ws As Object = ActiveWorkspace
     If TypeOf ws Is ResourceFileViewModel Then
      CType(ws, ResourceFileViewModel).SelectAllCommandHandler()
     End If
   End Select
  End Sub
#End Region

#Region " Status "
  Private _mainStatus As String = "Open or Create Project"
  Public Property MainStatus() As String
   Get
    Return _mainStatus
   End Get
   Set(ByVal value As String)
    _mainStatus = value
    Me.OnPropertyChanged("MainStatus")
   End Set
  End Property
#End Region

#Region " Locales "
  Private _availableLocales As List(Of CultureInfo)
  Public ReadOnly Property AvailableLocales() As List(Of CultureInfo)
   Get
    If _availableLocales Is Nothing Then
     _availableLocales = Common.Globals.AllLocales
    End If
    Return _availableLocales
   End Get
  End Property

  Private _selectedLocale As CultureInfo = New CultureInfo("en-US")
  Public Property SelectedLocale() As CultureInfo
   Get
    Return _selectedLocale
   End Get
   Set(ByVal value As CultureInfo)
    _selectedLocale = value
    Me.OnPropertyChanged("SelectedLocale")
   End Set
  End Property

  Private Sub SetLocale(locale As String)
   Dim ci As CultureInfo = AvailableLocales.FirstOrDefault(Function(x) x.Name = locale)
   If ci IsNot Nothing Then
    SelectedLocale = ci
   End If
  End Sub
#End Region

#Region " Settings "
  Private _ProjectSettings As Common.ProjectSettings
  Public Property ProjectSettings() As Common.ProjectSettings
   Get
    Return _ProjectSettings
   End Get
   Set(ByVal value As Common.ProjectSettings)
    _ProjectSettings = value
    Me.OnPropertyChanged("ProjectSettings")
    Try
     SelectedPackage = (From x In ProjectSettings.InstalledPackages Where x.PackageName = "Core" Select x)(0)
    Catch ex As Exception
    End Try
   End Set
  End Property

  Public Property TranslatorSettings As New Common.TranslatorSettings
  Private _settings As ViewModel.SettingsViewModel
  Public ReadOnly Property Settings As ViewModel.SettingsViewModel
   Get
    Return New ViewModel.SettingsViewModel(Me)
   End Get
  End Property
#End Region

#Region " Folder/File Tree "
  Private _treeContent As ObservableCollection(Of TreeViewItemViewModel)
  Public ReadOnly Property TreeContent As ObservableCollection(Of TreeViewItemViewModel)
   Get
    If _treeContent Is Nothing Then
     _treeContent = New ObservableCollection(Of TreeViewItemViewModel)
     If Not String.IsNullOrEmpty(_currentLocation) Then
      Dim fileCommands As New List(Of CommandViewModel)
      fileCommands.Add(New CommandViewModel("Open", OpenResourceFileCommand))
      Dim rf As ResourceFolderTreeViewModel
      If SelectedPackage.PackageName = "Core" Then
       rf = New ResourceFolderTreeViewModel(Nothing, "", ProjectSettings.CurrentSnapShot.ResourceFiles, fileCommands, New List(Of CommandViewModel))
      Else
       rf = New ResourceFolderTreeViewModel(Nothing, ProjectSettings.Location, SelectedPackage.Manifest.ResourceFiles.ToTree.FindCommonEntryPoint, ProjectSettings.CurrentSnapShot.ResourceFiles, fileCommands, New List(Of CommandViewModel))
      End If
      _treeContent = rf.Children
     End If
    End If
    Return _treeContent
   End Get
  End Property
#End Region

#Region " Subfolder/file Tree "
  Public ReadOnly Property HasSelectedPackage As Boolean
   Get
    Return CBool(SelectedPackage IsNot Nothing)
   End Get
  End Property

  Private _SelectedPackage As InstalledPackageViewModel
  Public Property SelectedPackage() As InstalledPackageViewModel
   Get
    Return _SelectedPackage
   End Get
   Set(ByVal value As InstalledPackageViewModel)
    _SelectedPackage = value
    _treeContent = Nothing
    Me.OnPropertyChanged("SelectedPackage")
    Me.OnPropertyChanged("TreeContent")
    Me.OnPropertyChanged("SubTreeContent")
    Me.OnPropertyChanged("HasSelectedPackage")
   End Set
  End Property

  Private Sub PackageSelected(sender As Object, args As PropertyChangedEventArgs)
   SelectedPackage = CType(sender, InstalledPackageViewModel)
   _subTreeContent = Nothing
   Me.OnPropertyChanged("SubTreeContent")
   Me.OnPropertyChanged("HasSelectedPackage")
  End Sub

  Private _subTreeContent As ObservableCollection(Of TreeViewItemViewModel)
  Public ReadOnly Property SubTreeContent As ObservableCollection(Of TreeViewItemViewModel)
   Get
    If _subTreeContent Is Nothing Then
     _subTreeContent = New ObservableCollection(Of TreeViewItemViewModel)
     If SelectedPackage IsNot Nothing Then
      Dim fileCommands As New List(Of CommandViewModel)
      fileCommands.Add(New CommandViewModel("Open", OpenResourceFileCommand))
      If SelectedPackage.PackageName = "Core" Then
       Dim f As New FolderViewModel(New IO.DirectoryInfo(_currentLocation), fileCommands, New List(Of CommandViewModel))
       f.IsExpanded = True
       For Each fvm As FolderViewModel In f.Children
        _subTreeContent.Add(fvm)
       Next
      Else
       Dim tree As Common.TreeItem = SelectedPackage.Manifest.ResourceFiles.ToTree().FindCommonEntryPoint
       Dim f As New FolderViewModel(Me.ProjectSettings.Location, tree, fileCommands, New List(Of CommandViewModel))
      End If
     End If
    End If
    Return _subTreeContent
   End Get
  End Property
#End Region

#Region " Workspaces "
  Private _workspaces As ObservableCollection(Of WorkspaceViewModel)

  ''' <summary>
  ''' Returns the collection of available workspaces to display.
  ''' A 'workspace' is a ViewModel that can request to be closed.
  ''' </summary>
  Public ReadOnly Property Workspaces() As ObservableCollection(Of WorkspaceViewModel)
   Get
    If _workspaces Is Nothing Then
     _workspaces = New ObservableCollection(Of WorkspaceViewModel)()
     AddHandler _workspaces.CollectionChanged, AddressOf OnWorkspacesChanged
    End If
    Return _workspaces
   End Get
  End Property

  Private Sub OnWorkspacesChanged(ByVal sender As Object, ByVal e As NotifyCollectionChangedEventArgs)
   If e.NewItems IsNot Nothing AndAlso e.NewItems.Count <> 0 Then
    For Each workspace As WorkspaceViewModel In e.NewItems
     AddHandler workspace.RequestClose, AddressOf OnWorkspaceRequestClose
    Next workspace
   End If

   If e.OldItems IsNot Nothing AndAlso e.OldItems.Count <> 0 Then
    For Each workspace As WorkspaceViewModel In e.OldItems
     RemoveHandler workspace.RequestClose, AddressOf OnWorkspaceRequestClose
    Next workspace
   End If
  End Sub

  Private Sub OnWorkspaceRequestClose(ByVal sender As Object, ByVal e As EventArgs)
   Dim workspace As WorkspaceViewModel = TryCast(sender, WorkspaceViewModel)
   workspace.Dispose()
   Me.Workspaces.Remove(workspace)
  End Sub

  Public Sub SetActiveWorkspace(ByVal workspace As WorkspaceViewModel)
   Debug.Assert(Me.Workspaces.Contains(workspace))

   Dim collectionView As ICollectionView = CollectionViewSource.GetDefaultView(Me.Workspaces)
   If collectionView IsNot Nothing Then
    collectionView.MoveCurrentTo(workspace)
   End If

  End Sub

  Public ReadOnly Property ActiveWorkspace As WorkspaceViewModel
   Get
    If Me.Workspaces.Count = 0 Or SelectedTabIndex = -1 Then Return Nothing
    Return Me.Workspaces(SelectedTabIndex)
   End Get
  End Property

  Private _selectedTabIndex As Integer = -1
  Public Property SelectedTabIndex() As Integer
   Get
    Return _selectedTabIndex
   End Get
   Set(ByVal value As Integer)
    _selectedTabIndex = value
    Me.OnPropertyChanged("ActiveResourceCollection")
   End Set
  End Property
#End Region

#Region " Resource Files "
  Public ReadOnly Property ActiveResourceCollection As ResourceCollectionViewModel
   Get
    If ActiveWorkspace Is Nothing Then Return Nothing
    If TypeOf ActiveWorkspace Is ResourceCollectionViewModel Then
     Return CType(ActiveWorkspace, ResourceCollectionViewModel)
    End If
    Return Nothing
   End Get
  End Property

  Private _openResourceFileCommand As RelayCommand
  Public ReadOnly Property OpenResourceFileCommand As RelayCommand
   Get
    If _openResourceFileCommand Is Nothing Then
     _openResourceFileCommand = New RelayCommand(Sub(param) Me.OpenResourceFileClicked(param))
    End If
    Return _openResourceFileCommand
   End Get
  End Property

  Protected Sub OpenResourceFileClicked(param As Object)
   IsBusy = True
   BusyMessage = "Opening Resource File"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf OpenResourceFile
   AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf OpenResourceFileCompleted
   Dim resfile As String = CType(param, String)
   If Not resfile.StartsWith(ProjectSettings.Location) Then resfile = ProjectSettings.Location & resfile
   Dim p As New Common.ParameterList(Me)
   p.Params.Add(resfile)
   _backgroundWorker.RunWorkerAsync(p)
  End Sub

  Private Sub OpenResourceFile(sender As Object, e As DoWorkEventArgs)
   Dim resource As String = CType(e.Argument, Common.ParameterList).Params(0)
   If resource.ToLower.EndsWith(".resx") Then
    Dim resFilename As String = resource.Substring(resource.LastIndexOf("\"c) + 1)
    Dim result As New OpenResourceFileResult
    result.mustAdd = False
    result.workspace = CType(Me.Workspaces.FirstOrDefault(Function(vm) vm.ID = resFilename), ResourceFileViewModel)
    If result.workspace Is Nothing Then
     result.workspace = New ResourceFileViewModel(CType(e.Argument, Common.ParameterList))
     result.mustAdd = True
    End If
    e.Result = result
   End If
  End Sub
  Private Structure OpenResourceFileResult
   Public workspace As ResourceFileViewModel
   Public mustAdd As Boolean
  End Structure

  Private Sub OpenResourceFileCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   Dim result As OpenResourceFileResult = CType(e.Result, Global.DotNetNuke.Translator.ViewModel.MainWindowViewModel.OpenResourceFileResult)
   If result.mustAdd Then
    Workspaces.Add(result.workspace)
   End If
   SetActiveWorkspace(result.workspace)
   IsBusy = False
  End Sub
#End Region

#Region " Recent Locations Stuff "
  Public Property RecentLocations As New Data.RecentLocations
  Private _recentLocationsList As ObservableCollection(Of RecentLocationViewModel)
  Public ReadOnly Property RecentLocationsList As ObservableCollection(Of RecentLocationViewModel)
   Get
    If _recentLocationsList Is Nothing Then
     _recentLocationsList = New ObservableCollection(Of RecentLocationViewModel)
     For Each kvp As KeyValuePair(Of Integer, Data.RecentLocation) In RecentLocations.RecentLocations
      _recentLocationsList.Add(New RecentLocationViewModel(kvp.Value, OpenRecentCommand))
     Next
    End If
    Return _recentLocationsList
   End Get
  End Property
#End Region

#Region " Opening/Saving "
  Private _openCommand As RelayCommand
  Public ReadOnly Property OpenCommand As RelayCommand
   Get
    If _openCommand Is Nothing Then
     _openCommand = New RelayCommand(Sub(param) Me.OpenClicked())
    End If
    Return _openCommand
   End Get
  End Property

  Private _openRecentCommand As RelayCommand
  Public ReadOnly Property OpenRecentCommand As RelayCommand
   Get
    If _openRecentCommand Is Nothing Then
     _openRecentCommand = New RelayCommand(Sub(param) Me.OpenRecentClicked(param))
    End If
    Return _openRecentCommand
   End Get
  End Property

  Protected Sub OpenClicked()

   Dim fbd As New System.Windows.Forms.OpenFileDialog With {.CheckFileExists = True, .DefaultExt = ".trproj", .Filter = "translation project files (*.trproj)|*.trproj"}
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    If IO.File.Exists(fbd.FileName) Then
     MainStatus = "Opening " & fbd.FileName
     Open(fbd.FileName)
    End If
   End If

  End Sub

  Protected Sub OpenRecentClicked(param As Object)
   Dim projectSettingsFile As String = CType(param, String)
   If Not IO.File.Exists(projectSettingsFile) Then
    MessageBox.Show("File no longer exists", "Selection Error", MessageBoxButton.OK)
    _RecentLocations.Remove(projectSettingsFile)
    MyBase.OnPropertyChanged("RecentLocationsList")
    Exit Sub
   End If
   Open(projectSettingsFile)
  End Sub

  Public Sub Open(location As String)

   If IO.File.Exists(location) Then ' it's a settings file
    _RecentLocations.Push(location)
    _recentLocationsList = Nothing ' reset
   ElseIf IO.Directory.Exists(location) Then ' it's a DNN directory
   End If

   IsBusy = True
   BusyMessage = "Opening Location"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf OpenLocation
   AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf OpenLocationCompleted
   _backgroundWorker.RunWorkerAsync(location)

  End Sub

  Private Sub OpenLocation(sender As Object, e As DoWorkEventArgs)

   Dim location As String = CType(e.Argument, String)
   ProjectSettings = New Common.ProjectSettings(location)
   e.Result = ProjectSettings

  End Sub

  Private Sub OpenLocationCompleted(sender As Object, e As RunWorkerCompletedEventArgs)

   Dim result As Common.ProjectSettings = CType(e.Result, Common.ProjectSettings)
   Me.ProjectSettings = result
   If ProjectSettings.TargetLocale <> "" Then
    SetLocale(ProjectSettings.TargetLocale)
   End If
   _treeContent = Nothing
   _currentLocation = ProjectSettings.Location
   MainStatus = "Location: " & ProjectSettings.Location

   MyBase.OnPropertyChanged("RecentLocationsList")
   MyBase.OnPropertyChanged("TreeContent")
   For Each ipvm As InstalledPackageViewModel In ProjectSettings.InstalledPackages
    AddHandler ipvm.PropertyChanged, AddressOf PackageSelected
   Next
   IsBusy = False

  End Sub

  Private _saveCommand As RelayCommand
  Public ReadOnly Property SaveCommand As RelayCommand
   Get
    If _saveCommand Is Nothing Then
     _saveCommand = New RelayCommand(Sub(param)
                                      Me.ProjectSettings.Save()
                                     End Sub,
                                     Function()
                                      Return CBool(Me.ProjectSettings IsNot Nothing AndAlso Me.ProjectSettings.SettingsFileName <> "")
                                     End Function)
    End If
    Return _saveCommand
   End Get
  End Property

  Private _SaveAsCommand As RelayCommand
  Public ReadOnly Property SaveAsCommand As RelayCommand
   Get
    If _SaveAsCommand Is Nothing Then
     _SaveAsCommand = New RelayCommand(Sub(param) Me.SaveAsClicked())
    End If
    Return _SaveAsCommand
   End Get
  End Property

  Protected Sub SaveAsClicked()

   Dim fbd As New System.Windows.Forms.SaveFileDialog With {.Filter = "translation project files (*.trproj)|*.trproj", .OverwritePrompt = True, .AddExtension = True, .CheckPathExists = True, .DefaultExt = ".trproj"}
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    Me.ProjectSettings.Save(fbd.FileName)
    _RecentLocations.Push(fbd.FileName)
    _recentLocationsList = Nothing ' reset
   End If

  End Sub

  Private _NewCommand As RelayCommand
  Public ReadOnly Property NewCommand As RelayCommand
   Get
    If _NewCommand Is Nothing Then
     _NewCommand = New RelayCommand(Sub(param) Me.NewClicked())
    End If
    Return _NewCommand
   End Get
  End Property

  Protected Sub NewClicked()

   Dim fbd As New System.Windows.Forms.FolderBrowserDialog
   fbd.Description = "Browse for target DNN installation"
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    If IO.Directory.Exists(fbd.SelectedPath) Then
     MainStatus = "Creating New Project at " & fbd.SelectedPath
     If IO.File.Exists(fbd.SelectedPath & "\web.config") AndAlso IO.File.Exists(fbd.SelectedPath & "\bin\dotnetnuke.dll") Then
      Open(fbd.SelectedPath)
     Else
      MessageBox.Show("You must select the root of a DotNetNuke installation", "Selection Error", MessageBoxButton.OK)
     End If
    End If
   End If

  End Sub
#End Region

#Region " Exiting "
  Public Sub ExitApplication(sender As Object, e As System.EventArgs) Handles Me.RequestClose
   RecentLocations.Save()
   Try
    ProjectSettings.Save()
   Catch ex As Exception
   End Try
  End Sub
#End Region

#Region " Bing "
  Public Property Bing As New Common.Bing.Bing(Me.TranslatorSettings)
#End Region

#Region " Language Packs "
  Private _savePack As RelayCommand
  Public ReadOnly Property SavePackCommand As RelayCommand
   Get
    If _savePack Is Nothing Then
     _savePack = New RelayCommand(Sub() Me.SavePackClicked(), Function() HasSelectedPackage)
    End If
    Return _savePack
   End Get
  End Property

  Protected Sub SavePackClicked()

   Dim start As String = _currentLocation
   Dim name As String = "DotNetNuke"
   Dim manifest As String = ""
   Dim version As String = Globals.NormalizeVersion(ProjectSettings.DotNetNukeVersion)
   If SelectedPackage IsNot Nothing AndAlso SelectedPackage.PackageName <> "Core" Then
    start &= "DesktopModules\" & SelectedPackage.FolderName
    name = SelectedPackage.FriendlyName
    version = Globals.NormalizeVersion(SelectedPackage.Version)
    manifest = SelectedPackage.ManifestXml
   End If
   Dim defaultPackname As String = name
   If defaultPackname = "DotNetNuke" Then
    Dim res As MsgBoxResult = MsgBox("Do you wish to generate a ""Full Core"" pack? If you select No only the core resources will be packed. If you select Yes, then all installed modules will be packed as well.", MsgBoxStyle.YesNoCancel, "Language Pack Generation")
    Select Case res
     Case MsgBoxResult.Yes
      defaultPackname = "DotNetNuke." & SelectedPackage.FriendlyName & ".Full"
      name = "Full"
     Case MsgBoxResult.No
      defaultPackname = "DotNetNuke.Core"
      name = "Core"
     Case MsgBoxResult.Cancel
      Exit Sub
    End Select
   End If

   Dim dlg As New Microsoft.Win32.SaveFileDialog()
   dlg.FileName = String.Format("ResourcePack.{0}.{1}.{2}.zip", defaultPackname, version, SelectedLocale.Name)
   dlg.DefaultExt = ".zip"
   dlg.Filter = "Language Packs (.zip)|*.zip"
   Dim result? As Boolean = dlg.ShowDialog()
   If result = True Then
    Dim filename As String = dlg.FileName
    IsBusy = True
    BusyMessage = "Creating Language Pack"
    _backgroundWorker = New BackgroundWorker
    AddHandler _backgroundWorker.DoWork, AddressOf SavePackFile
    AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf SavePackFileCompleted
    Dim p As New Common.ParameterList(Me)
    p.Params.Add(name) '0
    p.Params.Add(filename) '1
    _backgroundWorker.RunWorkerAsync(p)
   End If

  End Sub

  Private Sub SavePackFile(sender As Object, e As DoWorkEventArgs)
   Dim p As Common.ParameterList = CType(e.Argument, Common.ParameterList)

   'Dim pack As New Services.Packing.LanguagePack(p.Params(0), p.Params(1), p.Params(2), p.Params(3), p.Params(4), p.Params(5), p.Params(6), p.Params(7), p.Params(8), p.Params(9))
   Dim this As MainWindowViewModel = CType(p.ParentWindow, MainWindowViewModel)
   Dim pack As New Services.Packing.LanguagePack(this.TranslatorSettings, this.ProjectSettings, p.Params(0), this.SelectedLocale)
   pack.Save(p.Params(1))
  End Sub

  Private Sub SavePackFileCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   IsBusy = False
  End Sub
#End Region

#Region " Snapshots "
  Private _saveSnapshot As RelayCommand
  Public ReadOnly Property SaveSnapshotCommand As RelayCommand
   Get
    If _saveSnapshot Is Nothing Then
     _saveSnapshot = New RelayCommand(Sub() Me.SaveSnapshotClicked(), Function() HasSelectedPackage)
    End If
    Return _saveSnapshot
   End Get
  End Property

  'Protected Function SaveSnapshotEnabled() As Boolean
  ' Return HasSelectedPackage
  'End Function

  Protected Sub SaveSnapshotClicked()

   Dim start As String = _currentLocation
   Dim name As String = "DotNetNuke"
   Dim version As String = ProjectSettings.DotNetNukeVersion
   If SelectedPackage IsNot Nothing AndAlso SelectedPackage.PackageName <> "Core" Then
    start &= "DesktopModules\" & SelectedPackage.FolderName
    name = SelectedPackage.FriendlyName
    version = SelectedPackage.Version
   End If

   Dim dlg As New Microsoft.Win32.SaveFileDialog()
   dlg.FileName = String.Format("{0}_{1}", name, version)
   dlg.DefaultExt = ".tss"
   dlg.Filter = "Translator Snapshots (.tss)|*.tss"
   Dim result? As Boolean = dlg.ShowDialog()
   If result = True Then
    Dim filename As String = dlg.FileName
    IsBusy = True
    BusyMessage = "Saving Snapshot"
    _backgroundWorker = New BackgroundWorker
    AddHandler _backgroundWorker.DoWork, AddressOf SaveSnapshotFile
    AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf SaveSnapshotFileCompleted
    Dim p As New Common.ParameterList(Me)
    p.Params.Add(_currentLocation)
    p.Params.Add(start)
    p.Params.Add(filename)
    _backgroundWorker.RunWorkerAsync(p)
   End If

  End Sub

  Private Sub SaveSnapshotFile(sender As Object, e As DoWorkEventArgs)
   Dim p As Common.ParameterList = CType(e.Argument, Common.ParameterList)
   Dim ss As New Snapshot(p.Params(0), p.Params(1))
   ss.WriteXml(p.Params(2), System.Data.XmlWriteMode.IgnoreSchema)
  End Sub

  Private Sub SaveSnapshotFileCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   IsBusy = False
  End Sub

  Private _compareSnapshot As RelayCommand
  Public ReadOnly Property CompareSnapshotCommand As RelayCommand
   Get
    If _compareSnapshot Is Nothing Then
     _compareSnapshot = New RelayCommand(Sub() Me.CompareSnapshotClicked(), Function() Me.CompareSnapshotEnabled)
    End If
    Return _compareSnapshot
   End Get
  End Property

  Protected Function CompareSnapshotEnabled() As Boolean
   Return CBool(_currentLocation <> "")
  End Function

  Protected Sub CompareSnapshotClicked()

   Dim start As String = _currentLocation
   Dim name As String = "DotNetNuke"
   Dim version As String = ProjectSettings.DotNetNukeVersion
   If SelectedPackage IsNot Nothing AndAlso SelectedPackage.PackageName <> "Core" Then
    start &= "DesktopModules\" & SelectedPackage.FolderName
    name = SelectedPackage.FriendlyName
    version = SelectedPackage.Version
   End If

   Dim dlg As New Microsoft.Win32.OpenFileDialog()
   dlg.DefaultExt = ".tss"
   dlg.Filter = "Translator Snapshots (.tss)|*.tss"
   Dim result? As Boolean = dlg.ShowDialog()
   If result = True Then
    Try
     ' check if it isn't already open
     For Each ws As WorkspaceViewModel In Workspaces
      If ws.ID = String.Format("{0}_{1}_{2}", start, dlg.FileName, SelectedLocale.Name) Then
       SetActiveWorkspace(ws)
       Exit Sub
      End If
     Next
     IsBusy = True
     BusyMessage = "Comparing Snapshot"
     _backgroundWorker = New BackgroundWorker
     AddHandler _backgroundWorker.DoWork, AddressOf CompareSnapshotFile
     AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf CompareSnapshotFileCompleted
     Dim p As String() = {start, dlg.FileName, SelectedLocale.Name, name}
     _backgroundWorker.RunWorkerAsync(p)
    Catch ex As Exception
    End Try
   End If

  End Sub

  Private Sub CompareSnapshotFile(sender As Object, e As DoWorkEventArgs)
   Dim p As String() = CType(e.Argument, String())
   Dim ssc As New SnapshotComparisonViewModel(Me, p(0), p(1), p(2))
   ssc.DisplayName = p(3)
   e.Result = ssc
  End Sub

  Private Sub CompareSnapshotFileCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   Dim ssc As SnapshotComparisonViewModel = CType(e.Result, SnapshotComparisonViewModel)
   ssc.MainWindow = Me
   Me.Workspaces.Add(ssc)
   SetActiveWorkspace(ssc)
   IsBusy = False
  End Sub
#End Region

#Region " Search "
  Private _searchString As String = ""
  Public Property SearchString() As String
   Get
    Return _searchString
   End Get
   Set(ByVal value As String)
    _searchString = value
    Me.OnPropertyChanged("SearchString")
   End Set
  End Property

  Private _search As RelayCommand
  Public ReadOnly Property SearchCommand As RelayCommand
   Get
    If _search Is Nothing Then
     _search = New RelayCommand(Sub(param) Me.SearchClicked(param), Function(param) Me.SearchEnabled)
    End If
    Return _search
   End Get
  End Property

  Protected Function SearchEnabled() As Boolean
   Return CBool(_currentLocation <> "" And SearchString <> "")
  End Function

  Protected Sub SearchClicked(param As Object)

   Dim start As String = ""
   Dim name As String = "DotNetNuke"
   Dim version As String = ProjectSettings.DotNetNukeVersion
   If param IsNot Nothing AndAlso CType(param, String) <> "Core" AndAlso SelectedPackage IsNot Nothing Then
    start &= "DesktopModules\" & SelectedPackage.FolderName
    name = SelectedPackage.FriendlyName
    version = SelectedPackage.Version
   End If

   For Each ws As WorkspaceViewModel In Workspaces
    If ws.ID = String.Format("{0}_{1}_{2}", start, SearchString, SelectedLocale.Name) Then
     SetActiveWorkspace(ws)
     Exit Sub
    End If
   Next

   IsBusy = True
   BusyMessage = "Searching"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf SearchFile
   AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf SearchFileCompleted
   Dim p As New Common.ParameterList(Me)
   p.Params.Add(start)
   p.Params.Add(SearchString)
   _backgroundWorker.RunWorkerAsync(p)

  End Sub

  Private Sub SearchFile(sender As Object, e As DoWorkEventArgs)
   Dim svm As New SearchViewModel(CType(e.Argument, Common.ParameterList))
   e.Result = svm
  End Sub

  Private Sub SearchFileCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   Dim svm As SearchViewModel = CType(e.Result, SearchViewModel)
   svm.MainWindow = Me
   Workspaces.Add(svm)
   SetActiveWorkspace(svm)
   If svm.TooManyResults Then
    MsgBox("Too many results were found", MsgBoxStyle.Exclamation)
   End If
   IsBusy = False
  End Sub
#End Region

#Region " LE Connection "
  Private _LECompare As RelayCommand
  Public ReadOnly Property LECompareCommand As RelayCommand
   Get
    If _LECompare Is Nothing Then
     _LECompare = New RelayCommand(Sub() Me.LECompareClicked(), Function() Me.LECompareEnabled)
    End If
    Return _LECompare
   End Get
  End Property

  Protected Function LECompareEnabled() As Boolean
   Return HasSelectedPackage
  End Function

  Protected Sub LECompareClicked()

   If SelectedPackage Is Nothing Then Exit Sub

   Dim start As String = ""
   Dim version As String = ProjectSettings.DotNetNukeVersion
   Dim name As String = SelectedPackage.PackageName
   If name = "Core" Then
    Select Case ProjectSettings.DotNetNukeType
     Case "Professional Edition"
      name = "DNNPE"
     Case "Enterprise Edition"
      name = "DNNEE"
     Case Else
      name = "DNNCE"
    End Select
   ElseIf name = "DNN_HTML" Then
    If ProjectSettings.DotNetNukeType <> "Community Edition" Then name &= "_PRO"
   Else
    start &= "DesktopModules\" & SelectedPackage.FolderName
    version = SelectedPackage.Version
   End If

   For Each ws As WorkspaceViewModel In Workspaces
    If ws.ID = String.Format("Merge {0}", name) Then
     SetActiveWorkspace(ws)
     Exit Sub
    End If
   Next

   IsBusy = True
   BusyMessage = "LEComparing"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf LECompareFile
   AddHandler _backgroundWorker.RunWorkerCompleted, AddressOf LECompareFileCompleted
   Dim p As New Common.ParameterList(Me)
   p.Params.Add(name)
   p.Params.Add(Common.Globals.NormalizeVersion(version))
   p.Params.Add(start)
   _backgroundWorker.RunWorkerAsync(p)

  End Sub

  Private Sub LECompareFile(sender As Object, e As DoWorkEventArgs)
   Dim svm As New LEMergeViewModel(CType(e.Argument, Common.ParameterList))
   e.Result = svm
  End Sub

  Private Sub LECompareFileCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
   Dim svm As LEMergeViewModel = CType(e.Result, LEMergeViewModel)
   If Not svm.FailedCompare Then
    svm.MainWindow = Me
    Workspaces.Add(svm)
    SetActiveWorkspace(svm)
   End If
   IsBusy = False
  End Sub

  Public Sub SetMessage(msg As String)
   BusyMessage = msg
  End Sub
#End Region

 End Class
End Namespace
