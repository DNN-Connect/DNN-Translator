Namespace ViewModel
  Public Class FileViewModel
    Inherits TreeViewItemViewModel

    Private _file As IO.FileInfo

    Public Sub New(resourceFile As IO.FileInfo, fileCommands As List(Of CommandViewModel))
      MyBase.New(Nothing, True)
      _file = resourceFile
      For Each fc As CommandViewModel In fileCommands
        Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
      Next
    End Sub

    Public Sub New(parent As FolderViewModel, resourceFile As IO.FileInfo, fileCommands As List(Of CommandViewModel))
      MyBase.New(parent, True)
      _file = resourceFile
      For Each fc As CommandViewModel In fileCommands
        Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
      Next
    End Sub

    Public Overrides ReadOnly Property Key As String
      Get
        Return _file.FullName
      End Get
    End Property

    Public Overrides ReadOnly Property Name As String
      Get
        Return _file.Name
      End Get
    End Property

    Public Overrides ReadOnly Property Image As ImageSource
      Get
        Return CType(Application.Current.Resources("TextFile"), DrawingImage)
      End Get
    End Property

  End Class
End Namespace
