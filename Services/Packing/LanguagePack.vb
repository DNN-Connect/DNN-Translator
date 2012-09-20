Imports ICSharpCode.SharpZipLib.Zip
Imports DotNetNuke.Translator.Common

Namespace Services.Packing
 Public Class LanguagePack

  Public Property Manifest As PackManifest
  Private Property Settings As ProjectSettings
  Private Property LanguagePack As Byte()
  Private Property Locale As CultureInfo
  Private Property Copyright As String = ""
  Private Property AddedFiles As New List(Of String)

  Public Sub New(appSettings As TranslatorSettings, projectSettings As ProjectSettings, packageName As String, targetLocale As CultureInfo)

   Settings = projectSettings
   Locale = targetLocale
   If projectSettings.OverrideOwner Then
    Copyright = projectSettings.Copyright
   Else
    Copyright = appSettings.Copyright
   End If

   Manifest = New PackManifest(appSettings, projectSettings, Locale)

   Using LanguagePackStream As New IO.MemoryStream
    Using strmZipStream As New ZipOutputStream(LanguagePackStream)
     strmZipStream.SetLevel(9)
     For Each package As ViewModel.InstalledPackageViewModel In projectSettings.InstalledPackages
      If package.PackageName = packageName Or packageName = "Full" Then
       AddPackage(package, strmZipStream)
      End If
     Next
     Manifest.Clean()
     Dim manifestName As String = packageName & "_" & Locale.Name & ".dnn"
     manifestName = manifestName.Replace("/", "_").Replace("\", "_")
     Dim myZipEntry As New ZipEntry(manifestName)
     strmZipStream.PutNextEntry(myZipEntry)
     Dim fileData As Byte() = Globals.XmlToFormattedByteArray(Manifest)
     strmZipStream.Write(fileData, 0, fileData.Length)
     strmZipStream.Flush()
    End Using
    LanguagePack = LanguagePackStream.ToArray
   End Using

  End Sub

  Public Sub Save(path As String)

   Using strmZipFile As New IO.FileStream(path, IO.FileMode.Create)
    strmZipFile.Write(LanguagePack, 0, LanguagePack.Length)
   End Using

  End Sub

  Private Sub AddPackage(package As ViewModel.InstalledPackageViewModel, zipStream As ZipOutputStream)

   Dim packageNode As Xml.XmlNode = Manifest.CreatePackage(package)
   For Each fileKey As String In package.Manifest.ResourceFiles
    Dim targetFile As String = fileKey.Replace(".resx", "." & Locale.Name & ".resx")
    If Not AddedFiles.Contains(targetFile) AndAlso IO.File.Exists(Settings.Location & targetFile) Then
     AddedFiles.Add(targetFile)
     Dim resFile As New ResourceFile(targetFile, Settings.Location & targetFile)
     If resFile.Resources.Count > 0 Then
      Dim targetFileOriginalCase As String = Settings.CurrentSnapShot.ResFileOriginalCasings(fileKey).Replace(".resx", "." & Locale.Name & ".resx", StringComparison.InvariantCultureIgnoreCase)
      resFile.Copyright = Copyright
      resFile.Regenerate(True)
      Dim resFileData As Byte() = Globals.XmlToFormattedByteArray(resFile)
      Dim myZipEntry As New ZipEntry(targetFileOriginalCase)
      zipStream.PutNextEntry(myZipEntry)
      zipStream.Write(resFileData, 0, resFileData.Length)
      Manifest.AddResourceFile(packageNode, targetFileOriginalCase)
     End If
    End If
   Next

  End Sub

 End Class
End Namespace
