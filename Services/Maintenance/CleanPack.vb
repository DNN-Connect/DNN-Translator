Imports System.IO
Imports DotNetNuke.Translator.Common
Imports DotNetNuke.Translator.Common.Globals
Imports DotNetNuke.Translator.ViewModel

Namespace Services.Maintenance
  Public Class CleanPack
    Public Shared Sub Clean(projectSettings As ProjectSettings, package As InstalledPackageViewModel, targetLocale As CultureInfo)

      For Each fileKey As String In package.Manifest.ResourceFiles
        Dim enUs As New Common.ResourceFile(fileKey, projectSettings.Location & fileKey)
        Dim trans As New Common.ResourceFile(fileKey, GetResourceFilePath(projectSettings.Location, fileKey, targetLocale.Name))
        Dim keysToRemove As New List(Of String)
        For Each k As String In trans.Resources.Keys
          If Not enUs.Resources.ContainsKey(k) Then
            keysToRemove.Add(k)
          End If
        Next
        For Each k In keysToRemove
          trans.Resources.Remove(k)
        Next
        If trans.Resources.Count = 0 Then
          If File.Exists(trans.FilePath) Then
            File.Delete(trans.FilePath)
          End If
        Else
          trans.Save(projectSettings.DoNotUseLastModified)
        End If
      Next

    End Sub

  End Class
End Namespace