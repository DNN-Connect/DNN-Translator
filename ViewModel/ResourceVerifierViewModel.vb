Imports System.ComponentModel
Imports System.Collections.ObjectModel

Namespace ViewModel
 Public Class ResourceVerifierViewModel
  Inherits WorkspaceViewModel

#Region " Properties "
  Public Property MainWindow As MainWindowViewModel
  Public Property ObjectName As String
  Public Property ObjectVersion As String
  Public Property Start As String
  Public Property MissingKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property SameKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property TranslatedKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property ObsoleteKeys As New ObservableCollection(Of ResourceKeyViewModel)
  Public Property TotalOriginalResources As Integer = 0
  Public Property TotalTranslatedResources As Integer = 0
  Private _backgroundWorker As BackgroundWorker
#End Region

#Region " Constructor "
  Public Sub New(parameters As Common.ParameterList)
   Me.MainWindow = CType(parameters.ParentWindow, MainWindowViewModel)
   Me.ObjectName = parameters.Params(0)
   Me.ObjectVersion = parameters.Params(1)
   Me.Start = parameters.Params(2)
   Me.DisplayName = String.Format("Verify {0}", ObjectName)
   Me.ID = Me.DisplayName
   Compare()
  End Sub
#End Region

#Region " Comparing "
  Private Sub Compare()

   TotalOriginalResources = 0
   MissingKeys = New ObservableCollection(Of ResourceKeyViewModel)
   SameKeys = New ObservableCollection(Of ResourceKeyViewModel)
   TranslatedKeys = New ObservableCollection(Of ResourceKeyViewModel)
   ObsoleteKeys = New ObservableCollection(Of ResourceKeyViewModel)

   For Each resX As String In MainWindow.SelectedPackage.Manifest.ResourceFiles
    Dim enUs As New Common.ResourceFile(resX, MainWindow.ProjectSettings.Location & resX)
    Dim trans As New Common.ResourceFile(resX, Common.Globals.GetResourceFilePath(MainWindow.ProjectSettings.Location, resX, MainWindow.SelectedLocale.Name))
    For Each k As String In enUs.Resources.Keys
     TotalOriginalResources += 1
     If trans.Resources.ContainsKey(k) Then
      If trans.Resources(k).Value = enUs.Resources(k).Value Then
       SameKeys.Add(New ResourceKeyViewModel(enUs.Resources(k), trans.Resources(k).Value))
      Else
       TranslatedKeys.Add(New ResourceKeyViewModel(enUs.Resources(k), trans.Resources(k).Value))
      End If
     Else
      MissingKeys.Add(New ResourceKeyViewModel(enUs.Resources(k)))
     End If
    Next
    For Each k As String In trans.Resources.Keys
     TotalTranslatedResources += 1
     If Not enUs.Resources.ContainsKey(k) Then
      ObsoleteKeys.Add(New ResourceKeyViewModel(trans.Resources(k)))
     End If
    Next
   Next

   Me.OnPropertyChanged("SameKeys")
   Me.OnPropertyChanged("MissingKeys")
   Me.OnPropertyChanged("SameKeys")
   Me.OnPropertyChanged("TranslatedKeys")
   Me.OnPropertyChanged("ObsoleteKeys")
   Me.OnPropertyChanged("TotalOriginalResources")
   Me.OnPropertyChanged("TotalTranslatedResources")

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
    Case "translated"
     e.Result = ShowResourceKeys(what, "Translated Texts", TranslatedKeys)
    Case "obsolete"
     e.Result = ShowResourceKeys(what, "Obsolete Texts", ObsoleteKeys)
    Case "missing"
     e.Result = ShowResourceKeys(what, "Missing Texts", MissingKeys)
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
   resEdit.IsRemoteResourceCollection = False
   Return resEdit
  End Function
#End Region

 End Class
End Namespace
