Imports System.Text.RegularExpressions

Namespace ViewModel
  Public Class ResourceFolderTreeViewModel
    Inherits TreeViewItemViewModel

    Private _resourceFileCommands As List(Of CommandViewModel)
    Private _folderCommands As List(Of CommandViewModel)
    Private _key As String = ""
    Private _name As String = ""

    Public Sub New(parent As ResourceFolderTreeViewModel, dnnRoot As String, tree As Common.TreeItem, allResources As Dictionary(Of String, Common.ResourceFile), resourceFileCommands As List(Of CommandViewModel), folderCommands As List(Of CommandViewModel))
      MyBase.New(parent, False)

      _resourceFileCommands = resourceFileCommands
      _folderCommands = folderCommands
      _key = dnnRoot & tree.FullName
      _name = tree.Name

      For Each fc As CommandViewModel In folderCommands
        Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
      Next

      For Each child As Common.TreeItem In tree.Children
        If IO.Directory.Exists(dnnRoot & child.FullName) Then
          Try
            Me.Children.Add(New ResourceFolderTreeViewModel(Me, dnnRoot, child, allResources, resourceFileCommands, folderCommands))
          Catch ex As Exception
          End Try
        Else
          Me.Children.Add(New ResourceFileTreeViewModel(Me, allResources(child.FullName), resourceFileCommands))
        End If
      Next

    End Sub

    Public Sub New(parent As ResourceFolderTreeViewModel, start As String, allResources As Dictionary(Of String, Common.ResourceFile), resourceFileCommands As List(Of CommandViewModel), folderCommands As List(Of CommandViewModel))
      MyBase.New(parent, False)

      _resourceFileCommands = resourceFileCommands
      _folderCommands = folderCommands
      _key = start
      _name = _key.TrimEnd("\"c).Substring(_key.TrimEnd("\"c).LastIndexOf("\"c) + 1)

      For Each fc As CommandViewModel In folderCommands
        Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
      Next

      Dim subdirs As New List(Of String)
      For Each key As String In allResources.Keys
        If key.Length >= start.Length AndAlso key.Substring(0, start.Length) = start Then
          Dim relKey As String = key.Substring(start.Length)
          If relKey.IndexOf("\"c) > 0 Then ' it's still a directory underneath
            relKey = relKey.Substring(0, relKey.IndexOf("\"c))
            If Not String.IsNullOrEmpty(relKey) AndAlso Not subdirs.Contains(relKey) Then subdirs.Add(relKey)
          Else ' it's a resourcefile
            Me.Children.Add(New ResourceFileTreeViewModel(Me, allResources(key), resourceFileCommands))
          End If
        End If
      Next

      For Each subdir As String In subdirs
        Me.Children.Add(New ResourceFolderTreeViewModel(Me, start & subdir & "\", allResources, resourceFileCommands, folderCommands))
      Next

    End Sub

    Public Overrides ReadOnly Property Key As String
      Get
        Return _key
      End Get
    End Property

    Public Overrides ReadOnly Property Name As String
      Get
        Return _name
      End Get
    End Property

    Public Overrides ReadOnly Property Image As ImageSource
      Get
        Return CType(Application.Current.Resources("FolderClosed"), DrawingImage)
      End Get
    End Property

    Protected Overrides Sub LoadChildren()


    End Sub

  End Class
End Namespace
