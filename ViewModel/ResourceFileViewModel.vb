Imports System.Text.RegularExpressions
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Namespace ViewModel
 Public Class ResourceFileViewModel
  Inherits ResourceCollectionViewModel

  Private _originalResourceFile As String = ""
  Private _targetResourceFile As String = ""
  Private _fileName As String = ""
  Private _fileKey As String = ""
  Public Property OriginalFile As Common.ResourceFile
  Public Property TargetFile As Common.ResourceFile
  Public Property IsResourceFile As Boolean = True
  Private _backgroundWorker As BackgroundWorker

  Public Sub New(parameters As Common.ParameterObject)
   MyBase.New(CType(parameters.ParentWindow, MainWindowViewModel))
   _originalResourceFile = parameters.Params(0)
   _fileName = _originalResourceFile.Substring(_originalResourceFile.LastIndexOf("\"c) + 1)
   Me.DisplayName = _fileName
   Me.ID = _fileName
   _fileName = _fileName.Replace(".en-US.", ".") ' deal with DNN 6.2 template files
   _fileName = _fileName.Substring(0, _fileName.Length - 5) ' take off the resx
   _targetResourceFile = IO.Path.GetDirectoryName(_originalResourceFile) & "\" & _fileName & "." & TargetLocale.Name & ".resx" ' it won't normally find something if locale is empty
   _fileKey = _originalResourceFile.Replace(MainWindow.ProjectSettings.Location, "")
   _fileKey = _fileKey.TrimStart("\"c)

   ' Load the resource file
   OriginalFile = MainWindow.ProjectSettings.CurrentSnapShot.ResourceFiles(_fileKey)
   For Each Key As String In OriginalFile.Resources.Keys
    Dim rk As New ResourceKeyViewModel(OriginalFile.Resources(Key))
    AddHandler rk.PropertyChanged, AddressOf ResourceChanged
    ResourceKeys.Add(rk)
   Next
   ResourceKeys.Sort(Function(x) x.Key, ListSortDirection.Ascending)

   ' Load the target resource file
   TargetFile = New Common.ResourceFile(_fileKey, _targetResourceFile)
   For Each Key As String In TargetFile.Resources.Keys
    Dim k As String = Key
    Dim reskey As ResourceKeyViewModel = ResourceKeys.FirstOrDefault(Function(rk) rk.Key = k)
    If reskey IsNot Nothing Then
     reskey.TranslatedResource = TargetFile.Resources(Key)
    End If
   Next

   HasChanges = False

   ' check for alternate translations
   For Each f As String In IO.Directory.GetFiles(IO.Path.GetDirectoryName(_originalResourceFile), "*.resx")
    Dim fileName As String = f.Substring(f.LastIndexOf("\"c) + 1)
    Dim m As Match = Regex.Match(fileName, _fileName & "\.(\w{2,3}|\w{2,3}\-\w{2,3})\.resx")
    If m.Success Then
     _AvailableDiskResources.Add(m.Groups(1).Value)
    End If
   Next

  End Sub

  Protected Overrides Sub OnRequestClose()
   If HasChanges AndAlso TargetLocale IsNot Nothing Then
    Dim messageBoxText As String = "Do you want to save changes?"
    Dim caption As String = "Word Processor"
    Dim button As MessageBoxButton = MessageBoxButton.YesNoCancel
    Dim icon As MessageBoxImage = MessageBoxImage.Warning
    Dim result As MessageBoxResult = MessageBox.Show(messageBoxText, caption, button, icon)
    Select Case result
     Case MessageBoxResult.Yes
      TargetFile.Resources.Clear()
      For Each k As ResourceKeyViewModel In ResourceKeys
       If k.Changed Then
        TargetFile.SetResourceValue(k.Key, k.TargetValue, Now)
       Else
        TargetFile.SetResourceValue(k.Key, k.TargetValue, k.LastModified)
       End If
      Next
      TargetFile.Save()
     Case MessageBoxResult.No
      ' continue closing
     Case MessageBoxResult.Cancel
      Exit Sub
    End Select
   End If
   CloseMe()
  End Sub

  Private _AvailableDiskResources As New List(Of String)
  Public ReadOnly Property AvailableDiskResources As List(Of String)
   Get
    Return _AvailableDiskResources
   End Get
  End Property

#Region " Download Latesting "
  Private _DownloadLatestCommand As RelayCommand
  Public ReadOnly Property DownloadLatestCommand As RelayCommand
   Get
    If _DownloadLatestCommand Is Nothing Then
     _DownloadLatestCommand = New RelayCommand(Sub(param) DownloadLatestCommandHandler(param), Function(param) Me.DownloadLatestCommandEnabled)
    End If
    Return _DownloadLatestCommand
   End Get
  End Property

  Public Function DownloadLatestCommandEnabled() As Boolean
   Return True
  End Function

  Public Sub DownloadLatestCommandHandler(param As Object)
   MainWindow.IsBusy = True
   MainWindow.BusyMessage = "Downloading Resources"
   _backgroundWorker = New BackgroundWorker
   AddHandler _backgroundWorker.DoWork, AddressOf OpenResourceFile
   AddHandler _backgroundWorker.RunWorkerCompleted, Sub(sender As Object, e As RunWorkerCompletedEventArgs) MainWindow.IsBusy = False
   _backgroundWorker.RunWorkerAsync(param)
  End Sub

  Private Sub OpenResourceFile(sender As Object, e As DoWorkEventArgs)
   ' this is not foolproof!
   Dim objName As String = "Core"
   Dim objVersion As String = MainWindow.ProjectSettings.DotNetNukeVersion
   For Each p As InstalledPackageViewModel In MainWindow.ProjectSettings.InstalledPackages
    If _fileKey.StartsWith("DesktopModules\" & p.FolderName.Replace("/", "\")) Then
     objName = p.PackageName
     objVersion = p.Version
    End If
   Next
   ' now get the resources
   Dim service As New Common.LEService.LEService(MainWindow.ProjectSettings.ConnectionUrl, MainWindow.ProjectSettings.Username, MainWindow.ProjectSettings.Password)
   Try
    Dim remoteResources As List(Of Common.LEService.TextInfo) = service.GetResourceFile(objName, Common.Globals.NormalizeVersion(objVersion), MainWindow.ProjectSettings.MappedLocale, _fileKey)
    If remoteResources.Count = 0 Then
     MsgBox("No resources found", MsgBoxStyle.OkOnly)
    End If
    For Each ti As Common.LEService.TextInfo In remoteResources
     Dim k As String = ti.TextKey
     Dim reskey As ResourceKeyViewModel = ResourceKeys.FirstOrDefault(Function(rk) rk.Key = k)
     If reskey IsNot Nothing Then
      reskey.CompareValue = ti.Translation
     End If
    Next
   Catch ex As Exception
    If service.IsInError Then
     MsgBox(service.LastErrorMessage, MsgBoxStyle.OkOnly)
    Else
     MsgBox(ex.Message, MsgBoxStyle.Critical)
    End If
   End Try
  End Sub
#End Region

 End Class
End Namespace
