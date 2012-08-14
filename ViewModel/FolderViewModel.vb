Imports System.Text.RegularExpressions

Namespace ViewModel
 Public Class FolderViewModel
  Inherits TreeViewItemViewModel

  Private _directory As IO.DirectoryInfo
  Private _resourceFileCommands As List(Of CommandViewModel)
  Private _folderCommands As List(Of CommandViewModel)

  Public Sub New(directory As IO.DirectoryInfo, resourceFileCommands As List(Of CommandViewModel), folderCommands As List(Of CommandViewModel))
   MyBase.New(Nothing, True)
   _directory = directory
   _resourceFileCommands = resourceFileCommands
   _folderCommands = folderCommands
   For Each fc As CommandViewModel In folderCommands
    Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
   Next
  End Sub
  Public Sub New(parent As FolderViewModel, directory As IO.DirectoryInfo, resourceFileCommands As List(Of CommandViewModel), folderCommands As List(Of CommandViewModel))
   MyBase.New(parent, True)
   _directory = directory
   _resourceFileCommands = resourceFileCommands
   _folderCommands = folderCommands
   For Each fc As CommandViewModel In folderCommands
    Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
   Next
  End Sub

  Public Overrides ReadOnly Property Key As String
   Get
    Return _directory.FullName
   End Get
  End Property

  Public Overrides ReadOnly Property Name As String
   Get
    Return _directory.Name
   End Get
  End Property

  Public Overrides ReadOnly Property Image As System.Windows.Media.Imaging.BitmapImage
   Get
    Return New BitmapImage(New Uri("pack://application:,,,/Images/SmallIcon.png"))
   End Get
  End Property

  Protected Overrides Sub LoadChildren()

   For Each d As IO.DirectoryInfo In _directory.GetDirectories
    Children.Add(New FolderViewModel(Me, d, _resourceFileCommands, _folderCommands))
   Next
   For Each f As IO.FileInfo In _directory.GetFiles("*.resx")
    Dim m As Match = Regex.Match(f.Name, "\.(\w\w\-\w\w)\.resx")
    If m.Success AndAlso m.Groups(1).Value.ToLower <> "en-us" Then
     ' do nothing - it's a localized file
    Else
     Dim fvm As New FileViewModel(Me, f, _resourceFileCommands)
     fvm.IsExpanded = True
     Children.Add(fvm)
    End If
   Next

  End Sub

 End Class
End Namespace
