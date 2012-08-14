Imports System.Text.RegularExpressions

Namespace ViewModel
 Public Class ResourceFileTreeViewModel
  Inherits TreeViewItemViewModel

  Property ResourceFile As Common.ResourceFile
  Private _key As String = ""
  Private _name As String = ""

  Public Sub New(parent As ResourceFolderTreeViewModel, resource As Common.ResourceFile, resourceFileCommands As List(Of CommandViewModel))
   MyBase.New(parent, False)
   Me.ResourceFile = resource
   _key = ResourceFile.fileKey
   _name = _key.Substring(_key.LastIndexOf("\"c) + 1)
   For Each fc As CommandViewModel In resourceFileCommands
    Me.AddCommand(New CommandViewModel(fc.DisplayName, fc.Command, Key))
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

  Public Overrides ReadOnly Property Image As System.Windows.Media.Imaging.BitmapImage
   Get
    Return New BitmapImage(New Uri("pack://application:,,,/Images/SmallIcon.png"))
   End Get
  End Property

 End Class
End Namespace
