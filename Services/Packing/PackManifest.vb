Imports System.Xml
Imports DotNetNuke.Translator.Common

Namespace Services.Packing
 Public Class PackManifest
  Inherits XmlDocument

  Private Property Packages As XmlNode
  Private Property Locale As CultureInfo
  Private Property OwnerName As String = ""
  Private Property OwnerEmail As String = ""
  Private Property OwnerUrl As String = ""
  Private Property OwnerOrganization As String = ""
  Private Property License As String = ""

  Public Sub New(appSettings As TranslatorSettings, projectSettings As ProjectSettings, targetLocale As CultureInfo)
   MyBase.New()

   Me.Locale = targetLocale
   If projectSettings.OverrideOwner Then
    OwnerEmail = projectSettings.OwnerEmail
    OwnerName = projectSettings.OwnerName
    OwnerOrganization = projectSettings.OwnerOrganization
    OwnerUrl = projectSettings.OwnerUrl
    License = projectSettings.License
   Else
    OwnerEmail = appSettings.OwnerEmail
    OwnerName = appSettings.OwnerName
    OwnerOrganization = appSettings.OwnerOrganization
    OwnerUrl = appSettings.OwnerUrl
    License = appSettings.License
   End If

   Me.CreateXmlDeclaration("1.0", Nothing, Nothing)
   Dim dotnetnuke As XmlNode = Me.CreateElement("dotnetnuke")
   Me.AppendChild(dotnetnuke)
   dotnetnuke.AppendAttribute("type", "Package")
   dotnetnuke.AppendAttribute("version", "5.0")

   Packages = dotnetnuke.CreateAndAppendElement("packages")

  End Sub

  Public Function CreatePackage(originalPackage As ViewModel.InstalledPackageViewModel) As XmlNode

   Dim package As XmlNode = Packages.CreateAndAppendElement("package")
   package.AppendAttribute("name", String.Format("{0}_{1}", originalPackage.PackageName, Locale.Name))
   If originalPackage.PackageName = "Core" Then
    package.AppendAttribute("type", "CoreLanguagePack")
   Else
    package.AppendAttribute("type", "ExtensionLanguagePack")
   End If
   package.AppendAttribute("version", originalPackage.Version)
   package.CreateAndAppendElement("friendlyName", String.Format("{0} {1}", originalPackage.PackageName, Locale.NativeName))
   package.CreateAndAppendElement("description", String.Format("{1} language pack for {0}", originalPackage.PackageName, Locale.EnglishName))
   Dim owner As XmlNode = package.CreateAndAppendElement("owner")
   owner.CreateAndAppendElement("name", OwnerName)
   owner.CreateAndAppendElement("organization", OwnerOrganization)
   owner.CreateAndAppendElement("url", OwnerUrl)
   owner.CreateAndAppendElement("email", OwnerEmail)
   Dim lic As XmlNode = package.CreateAndAppendElement("license")
   lic.AppendChild(Me.CreateCDataSection(License))
   Dim components As XmlNode = package.CreateAndAppendElement("components")
   Dim component As XmlNode = components.CreateAndAppendElement("component")
   If originalPackage.PackageName = "Core" Then
    component.AppendAttribute("type", "CoreLanguage")
   Else
    component.AppendAttribute("type", "ExtensionLanguage")
   End If
   Dim languageFiles As XmlNode = component.CreateAndAppendElement("languageFiles")
   languageFiles.CreateAndAppendElement("code", Locale.Name)
   languageFiles.CreateAndAppendElement("displayName", Locale.NativeName)
   languageFiles.CreateAndAppendElement("basePath", "")
   Return languageFiles

  End Function

  Public Sub AddResourceFile(ByRef parent As XmlNode, filePath As String)
   Dim languageFile As XmlNode = parent.CreateAndAppendElement("languageFile")
   languageFile.CreateAndAppendElement("path", IO.Path.GetDirectoryName(filePath))
   languageFile.CreateAndAppendElement("name", IO.Path.GetFileName(filePath))
  End Sub

 End Class
End Namespace
