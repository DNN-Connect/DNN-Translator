Imports System.Data.SqlClient

Namespace ViewModel
 Public Class InstalledPackageViewModel
  Inherits ViewModelBase

  Public Property PackageName As String = ""
  Public Property FriendlyName As String = ""
  Public Property Version As String = ""
  Public Property FolderName As String = ""
  Public Property IsSystemPackage As Boolean = False
  Public Property Owner As String = ""
  Public Property ManifestXml As String = ""
  Public Property Manifest As Common.ModuleManifest

  Private _selected As Boolean = False
  Public Property Selected() As Boolean
   Get
    Return _selected
   End Get
   Set(ByVal value As Boolean)
    _selected = value
    Me.OnPropertyChanged("Selected")
   End Set
  End Property

  Public Sub New(packageName As String, friendlyName As String, version As String, folderName As String)
   Me.PackageName = packageName
   Me.FriendlyName = friendlyName
   Me.Version = version
   Me.FolderName = folderName
  End Sub

  Public Sub New(ir As SqlDataReader)
   MyBase.New()

   PackageName = GetString(ir, "Name")
   FriendlyName = GetString(ir, "FriendlyName")
   Version = GetString(ir, "Version")
   FolderName = GetString(ir, "ModuleFolderName")
   IsSystemPackage = GetBoolean(ir, "IsSystemPackage")
   Owner = GetString(ir, "Owner")
   ManifestXml = GetString(ir, "Manifest")

  End Sub

  Public Function LoadManifest(dnnBasePath As String) As Boolean
   If ManifestXml = "" Then Return False
   Manifest = New Common.ModuleManifest(ManifestXml, dnnBasePath)
   If Manifest.ResourceFiles.Count = 0 Then Return False
   Return True
  End Function

  Private Function GetString(ir As SqlDataReader, columnName As String) As String
   If ir.Item(columnName) Is DBNull.Value Then
    Return ""
   Else
    Return CStr(ir.Item(columnName))
   End If
  End Function

  Private Function GetBoolean(ir As SqlDataReader, columnName As String) As Boolean
   If ir.Item(columnName) Is DBNull.Value Then
    Return False
   Else
    Return CBool(ir.Item(columnName))
   End If
  End Function

 End Class
End Namespace
