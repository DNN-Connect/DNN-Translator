Imports System.Xml
Imports System.Text.RegularExpressions

Namespace Common
 Public Class ModuleManifest
  Inherits XmlDocument

  Public Property ResourceFiles As New List(Of String)
  Public Property ResourcePaths As New List(Of String)

  Public Sub New()
   MyBase.New()
  End Sub

  Public Sub New(objectManifest As String, dnnLocation As String)
   MyBase.New()
   MyBase.LoadXml(objectManifest)

   For Each f As XmlNode In Me.DocumentElement.SelectNodes("/package/components/component[@type='File']/files/file[substring(name, string-length(name) - 4)='.resx']")
    Dim fname As String = f.SelectSingleNode("name").InnerText.Replace("/", "\").ToLower
    If f.SelectSingleNode("path") IsNot Nothing Then
     fname = f.SelectSingleNode("path").InnerText.Replace("/", "\").ToLower & "\" & fname
    End If
    If f.ParentNode.SelectSingleNode("basePath") IsNot Nothing Then
     fname = f.ParentNode.SelectSingleNode("basePath").InnerText.Replace("/", "\").ToLower & "\" & fname
    End If
    fname = fname.Trim("\"c).Replace("\\", "\")
    If Not ResourceFiles.Contains(fname) Then ResourceFiles.Add(fname)
   Next

   For Each rf As XmlNode In Me.DocumentElement.SelectNodes("/package/components/component[@type='ResourceFile']/resourceFiles/resourceFile")
    Dim path As String = ""
    If rf.ParentNode.SelectSingleNode("basePath") Is Nothing Then
     If rf.ParentNode.ParentNode.ParentNode.SelectSingleNode("component[@type='Module']") IsNot Nothing Then
      path = rf.ParentNode.ParentNode.ParentNode.SelectSingleNode("component[@type='Module']/desktopModule/foldername").InnerText.Replace("/", "\").ToLower.Trim("\"c)
     End If
    Else
     path = rf.ParentNode.SelectSingleNode("basePath").InnerText.Replace("/", "\").ToLower.Trim("\"c)
    End If
    If path <> "" And path <> "DesktopModules\DNNCorp" And Not ResourcePaths.Contains(path) Then ResourcePaths.Add(path)
   Next

   If Not String.IsNullOrEmpty(dnnLocation) Then
    If Not dnnLocation.EndsWith("\"c) Then dnnLocation &= "\"
    For Each p As String In ResourcePaths
     ScanForResourceFiles(dnnLocation, p)
    Next
   End If

  End Sub

  Private Sub ScanForResourceFiles(basepath As String, path As String)

   If Not IO.Directory.Exists(basepath & path) Then Exit Sub
   Dim d As New IO.DirectoryInfo(basepath & path)
   For Each f As IO.FileInfo In d.GetFiles("*.resx")
    Dim m As Match = Regex.Match(f.Name, "\.(\w\w\-\w\w)\.resx")
    If m.Success AndAlso m.Groups(1).Value.ToLower <> "en-us" Then
     ' do nothing - it's a localized file
    Else
     Dim fname As String = path
     If fname <> "" Then fname &= "\"
     fname &= f.Name.ToLower().Replace("\\", "\")
     If Not ResourceFiles.Contains(fname) Then ResourceFiles.Add(fname)
    End If
   Next
   If path <> "" Then path &= "\"
   For Each subdir As IO.DirectoryInfo In d.GetDirectories
    ScanForResourceFiles(basepath, path & subdir.Name.ToLower)
   Next

  End Sub

 End Class
End Namespace
