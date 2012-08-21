Imports DotNetNuke.Translator.Common

Namespace Services.Packing
 Public Class LanguagePack

  Public Property Manifest As PackManifest

  Public Sub New(appSettings As TranslatorSettings, projectSettings As ProjectSettings, packageName As String)

   Dim locale As String = projectSettings.Locale.Name


  End Sub

  Public Sub Save(path As String)

  End Sub

 End Class
End Namespace
