Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Windows.Controls

Namespace Controls.FolderBrowser
 Public Class FolderViewModel
  Inherits ViewModelBase
  Private _isSelected As Boolean
  Private _isExpanded As Boolean
  Private _folderIcon As String

  Public Property Root() As BrowserViewModel
   Get
    Return m_Root
   End Get
   Set(value As BrowserViewModel)
    m_Root = value
   End Set
  End Property
  Private m_Root As BrowserViewModel

  Public Property FolderIcon() As String
   Get
    Return _folderIcon
   End Get
   Set(value As String)
    _folderIcon = value
    OnPropertyChanged("FolderIcon")
   End Set
  End Property

  Public Property FolderName() As String
   Get
    Return m_FolderName
   End Get
   Set(value As String)
    m_FolderName = value
   End Set
  End Property
  Private m_FolderName As String

  Public Property FolderPath() As String
   Get
    Return m_FolderPath
   End Get
   Set(value As String)
    m_FolderPath = value
   End Set
  End Property
  Private m_FolderPath As String

  Public Property Folders() As ObservableCollection(Of FolderViewModel)
   Get
    Return m_Folders
   End Get
   Set(value As ObservableCollection(Of FolderViewModel))
    m_Folders = value
   End Set
  End Property
  Private m_Folders As ObservableCollection(Of FolderViewModel)

  Public Property IsSelected() As Boolean
   Get
    Return _isSelected
   End Get
   Set(value As Boolean)
    If _isSelected <> value Then
     _isSelected = value

     OnPropertyChanged("IsSelected")

     If value Then
      Root.SelectedFolder = FolderPath
      'Default windows behaviour of expanding the selected folder
      IsExpanded = True
     End If
    End If
   End Set
  End Property

  Public Property IsExpanded() As Boolean
   Get
    Return _isExpanded
   End Get
   Set(value As Boolean)
    If _isExpanded <> value Then
     _isExpanded = value

     OnPropertyChanged("IsExpanded")

     If Not FolderName.Contains(":"c) Then
      'Folder icon change not applicable for drive(s)
      If _isExpanded Then
       FolderIcon = "Images\FolderBrowser\FolderOpen.png"
      Else
       FolderIcon = "Images\FolderBrowser\FolderClosed.png"
      End If
     End If

     LoadFolders()

    End If
   End Set
  End Property

  Private Sub LoadFolders()
   Try
    If Folders.Count > 0 Then
     Return
    End If

    Dim dirs As String() = Nothing

    Dim fullPath As String = Path.Combine(FolderPath, FolderName)

    If FolderName.Contains(":"c) Then
     'This is a drive
     fullPath = String.Concat(FolderName, "\")
    Else
     fullPath = FolderPath
    End If

    dirs = Directory.GetDirectories(fullPath)

    Folders.Clear()

    For Each dir As String In dirs
     Try
      Dim di As New DirectoryInfo(dir)
      ' create the sub-structure only if this is not a hidden directory
      If (di.Attributes And FileAttributes.Hidden) <> FileAttributes.Hidden Then
       Folders.Add(New FolderViewModel() With {
         .Root = Me.Root,
         .FolderName = Path.GetFileName(dir),
         .FolderPath = Path.GetFullPath(dir),
         .FolderIcon = "Images\FolderBrowser\FolderClosed.png"
       })
      End If
     Catch ae As UnauthorizedAccessException
      Console.WriteLine(ae.Message)
     Catch e As Exception
      Console.WriteLine(e.Message)
     End Try
    Next

    If FolderName.Contains(":") Then
     FolderIcon = "Images\FolderBrowser\HardDisk.ico"
    End If
   Catch ae As UnauthorizedAccessException
    Console.WriteLine(ae.Message)
   Catch ie As IOException
    Console.WriteLine(ie.Message)
   End Try
  End Sub

  Public Sub New()
   Folders = New ObservableCollection(Of FolderViewModel)()
  End Sub
 End Class
End Namespace
