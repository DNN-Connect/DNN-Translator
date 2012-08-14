Imports System.Text.RegularExpressions
Imports DotNetNuke.Translator.Common

Partial Class Snapshot

 Private _basePath As String = ""

 Public Property ResourceFiles As Dictionary(Of String, ResourceFile)

 Public Sub New(basepath As String, location As String)
  MyBase.New()
  Me.BeginInit()
  Me.InitClass()
  Dim schemaChangedHandler As Global.System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
  AddHandler MyBase.Tables.CollectionChanged, schemaChangedHandler
  AddHandler MyBase.Relations.CollectionChanged, schemaChangedHandler
  Me.EndInit()
  _basePath = basepath.TrimEnd("\"c)
  location = location.Substring(basepath.Length)
  location = location.Trim("\"c)
  If location <> "" And Not location.StartsWith("\"c) Then location = "\" & location
  ResourceFiles = New Dictionary(Of String, ResourceFile)
  AddResources(location)
 End Sub

 Private Sub AddResources(path As String)

  Try
   Dim directory As String = _basePath & path
   For Each f As String In IO.Directory.GetFiles(directory, "*.resx")
    Dim m As Match = Regex.Match(f, "\.(\w\w\-\w\w)\.resx")
    If m.Success AndAlso m.Groups(1).Value.ToLower <> "en-us" Then
     ' do nothing - it's a localized file
    Else
     Try
      Dim fileKey As String = path & "\" & IO.Path.GetFileName(f)
      fileKey = fileKey.TrimStart("\"c)
      Dim resFile As New ResourceFile(fileKey, f)
      ResourceFiles.Add(fileKey, resFile)
      For Each key As String In resFile.Resources.Keys
       Dim rr As ResourcesRow = Me.Resources.NewResourcesRow
       rr.FileKey = fileKey
       rr.ResourceKey = key
       rr.ResourceValue = System.Web.HttpUtility.HtmlDecode(resFile.Resources(key).Value)
       Me.Resources.AddResourcesRow(rr)
      Next
     Catch ex As Exception
      ' ignore files I can't read
     End Try
    End If
   Next

   For Each d As String In IO.Directory.GetDirectories(directory)
    d = d.Substring(d.LastIndexOf("\"c))
    AddResources(path & d)
   Next
  Catch ex As Exception

  End Try

 End Sub

End Class
