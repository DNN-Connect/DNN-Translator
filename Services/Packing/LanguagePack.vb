Imports System.IO
Imports System.IO.Compression
Imports DotNetNuke.Translator.Common
Imports DotNetNuke.Translator.Common.Globals

Namespace Services.Packing
  Public Class LanguagePack

    Public Property Manifest As PackManifest
    Private Property Settings As ProjectSettings
    Private Property LanguagePack As Byte()
    Private Property Locale As CultureInfo
    Private Property Copyright As String = ""
    Private Property AddedFiles As New List(Of String)

    Public Sub New(appSettings As TranslatorSettings, projectSettings As ProjectSettings, packageFriendlyName As String, targetLocale As CultureInfo)

      Settings = projectSettings
      Locale = targetLocale
      If projectSettings.OverrideOwner Then
        Copyright = projectSettings.Copyright
      Else
        Copyright = appSettings.Copyright
      End If

      Manifest = New PackManifest(appSettings, projectSettings, Locale)

      Using LanguagePackStream As New MemoryStream
        Using archive As New ZipArchive(LanguagePackStream, ZipArchiveMode.Create, False)
          For Each package As ViewModel.InstalledPackageViewModel In projectSettings.InstalledPackages
            If package.FriendlyName = packageFriendlyName Or packageFriendlyName = "Full" Then
              AddPackage(package, archive)
            End If
          Next
          Manifest.Clean()
          Dim manifestName As String = packageFriendlyName & "_" & Locale.Name & ".dnn"
          manifestName = manifestName.Replace("/", "_").Replace("\", "_")
          TranslatorLibrary.Compression.ZipXmlFileToStream(archive, Manifest, manifestName)
        End Using
        LanguagePack = LanguagePackStream.ToArray()
      End Using

    End Sub

    Public Sub Save(path As String)

      Using strmZipFile As New IO.FileStream(path, IO.FileMode.Create)
        strmZipFile.Write(LanguagePack, 0, LanguagePack.Length)
      End Using

    End Sub

    Private Sub AddPackage(package As ViewModel.InstalledPackageViewModel, archive As ZipArchive)

      Dim packageNode As Xml.XmlNode = Manifest.CreatePackage(package)
      For Each fileKey As String In package.Manifest.ResourceFiles
        Dim targetFile As String = GetLocalizedFilePath(fileKey, Locale.Name)
        If Not AddedFiles.Contains(targetFile) AndAlso IO.File.Exists(Settings.Location & targetFile) Then
          AddedFiles.Add(targetFile)

          Dim resFile As New ResourceFile(targetFile, Settings.Location & targetFile)
          If resFile.Resources.Count > 0 Then
            Dim targetFileOriginalCase As String = GetLocalizedFilePath(Settings.CurrentSnapShot.ResFileOriginalCasings(fileKey), Locale.Name)
            resFile.Copyright = Copyright
            resFile.Regenerate(True)
            TranslatorLibrary.Compression.ZipXmlFileToStream(archive, resFile, targetFileOriginalCase)
            Manifest.AddResourceFile(packageNode, targetFileOriginalCase)
          End If
        End If
      Next

    End Sub

  End Class
End Namespace
