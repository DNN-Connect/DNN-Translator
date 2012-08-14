Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.IO
Imports DotNetNuke.Translator.Common

Namespace Controls.FolderBrowser
 Public Class BrowserViewModel
  Inherits ViewModelBase
  Private _selectedFolder As String
  Private _expanding As Boolean = False

  Public Property SelectedFolder() As String
   Get
    Return _selectedFolder
   End Get
   Set(value As String)
    _selectedFolder = value
    OnPropertyChanged("SelectedFolder")
    OnSelectedFolderChanged()
   End Set
  End Property

  Public Property Folders() As ObservableCollection(Of FolderViewModel)
   Get
    Return m_Folders
   End Get
   Set(value As ObservableCollection(Of FolderViewModel))
    m_Folders = value
   End Set
  End Property
  Private m_Folders As ObservableCollection(Of FolderViewModel)

  Public ReadOnly Property FolderSelectedCommand() As DelegateCommand(Of Object)
   Get
    Return New DelegateCommand(Of Object)(Function(it) InlineAssignHelper(SelectedFolder, Environment.GetFolderPath(CType(it, Environment.SpecialFolder))))
   End Get
  End Property


  Public Sub New()
   Folders = New ObservableCollection(Of FolderViewModel)()
   Environment.GetLogicalDrives().ToList().ForEach(Sub(it) Folders.Add(New FolderViewModel() With {
     .Root = Me,
     .FolderPath = it.TrimEnd("\"c),
     .FolderName = it.TrimEnd("\"c),
     .FolderIcon = "Images\FolderBrowser\HardDisk.ico"
   }))
  End Sub

  Private Sub OnSelectedFolderChanged()
   If Not _expanding Then
    Try
     _expanding = True
     Dim child As FolderViewModel = Expand(Folders, SelectedFolder)
     child.IsSelected = True
    Catch
    Finally
     _expanding = False
    End Try
   End If
  End Sub

  Private Function Expand(childFolders As ObservableCollection(Of FolderViewModel), path As String) As FolderViewModel
   If [String].IsNullOrEmpty(path) OrElse childFolders.Count = 0 Then
    Return Nothing
   End If

   Dim folderName As String = path
   If path.Contains("/"c) OrElse path.Contains("\"c) Then
    Dim idx As Integer = path.IndexOfAny(New Char() {"/"c, "\"c})
    folderName = path.Substring(0, idx)
    path = path.Substring(idx + 1)
   Else
    path = Nothing
   End If

   Dim results = childFolders.Where(Function(folder)
                                     Return CBool(folder.FolderName = folderName)
                                    End Function)
   If results IsNot Nothing AndAlso results.Count() > 0 Then
    Dim fvm As FolderViewModel = results.First()
    fvm.IsExpanded = True

    Dim retVal = Expand(fvm.Folders, path)
    If retVal IsNot Nothing Then
     Return retVal
    Else
     Return fvm
    End If
   End If

   Return Nothing
  End Function

  Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
   target = value
   Return value
  End Function


 End Class
End Namespace
