Imports System.IO.IsolatedStorage
Imports System.IO

Namespace Common
 Public Class TranslatorSettings
  Inherits TranslatorData

  Private Shadows Const SettingsFilename As String = "TranslatorSettings.xml"

        Public Property TranslationProvider As String = "none"
        Public Property AzureApiKey As String = ""
        Public Property AzureRegion As String = ""
        Public Property AzureLanguages As String = ""
        Public Property AzureLastLanguagesRetrieve As Date = Date.MinValue
        Public Property BingClientID As String = ""
        Public Property BingClientSecret As String = ""
        Public Property BingLanguages As String = ""
        Public Property BingLastLanguagesRetrieve As Date = Date.MinValue
        Public Property GoogleApiKey As String = ""
  Public Property DefaultTargetLocale As String = ""
  Public Property MaximumKeys As Integer = 100
  Public Property OwnerName As String = ""
  Public Property OwnerEmail As String = ""
  Public Property OwnerUrl As String = ""
  Public Property OwnerOrganization As String = ""
  Public Property License As String = ""
  Public Property Copyright As String = ""
  Public Property DefaultDictionary As String = ""
  Public Property LastUsedDir As String = ""

  Public Sub New()
   MyBase.New()

   Dim isoStore As IsolatedStorageFile = IsolatedStorageFile.GetStore(IsolatedStorageScope.User Or IsolatedStorageScope.Assembly Or IsolatedStorageScope.Domain, Nothing, Nothing)
   Dim recentFiles As String() = isoStore.GetFileNames(SettingsFilename)
   If recentFiles.Length > 0 Then
    Dim i As Integer = 1
    Using strIn As New IO.StreamReader(New IsolatedStorageFileStream(SettingsFilename, FileMode.Open, isoStore))
     Try
      Me.ReadXml(strIn, System.Data.XmlReadMode.IgnoreSchema)
     Catch ex As Exception
     End Try
    End Using
   End If

   ReadSettingValue("TranslationProvider", TranslationProvider)
   ReadSettingValue("BingClientID", BingClientID)
   ReadSettingValue("BingClientSecret", BingClientSecret)
            ReadSettingValue("BingLanguages", BingLanguages)
            ReadSettingValue("BingLastLanguagesRetrieve", BingLastLanguagesRetrieve)
            ReadSettingValue("AzureLanguages", AzureLanguages)
            ReadSettingValue("AzureLastLanguagesRetrieve", AzureLastLanguagesRetrieve)
            ReadSettingValue("AzureApiKey", AzureApiKey)
            ReadSettingValue("AzureRegion", AzureRegion)
            ReadSettingValue("GoogleApiKey", GoogleApiKey)
            ReadSettingValue("DefaultTargetLocale", DefaultTargetLocale)
            ReadSettingValue("OwnerName", OwnerName)
   ReadSettingValue("OwnerEmail", OwnerEmail)
   ReadSettingValue("OwnerUrl", OwnerUrl)
   ReadSettingValue("OwnerOrganization", OwnerOrganization)
   ReadSettingValue("License", License)
   ReadSettingValue("Copyright", Copyright)
   ReadSettingValue("DefaultDictionary", DefaultDictionary)
   ReadSettingValue("LastUsedDir", LastUsedDir)

   If DefaultDictionary = "" Then
    DefaultDictionary = Common.Globals.TranslatorDocFolder & "\MyDictionary.dic"
   End If

  End Sub

  Public Overrides Sub Save()

   Setting("TranslationProvider", False) = TranslationProvider
   Setting("BingClientID", False) = BingClientID
   Setting("BingClientSecret", False) = BingClientSecret
   Setting("BingLanguages", False) = BingLanguages
            Setting("BingLastLanguagesRetrieve", False) = BingLastLanguagesRetrieve.ToString("yyyy-MM-dd HH:mm:ss")
            Setting("AzureLanguages", False) = AzureLanguages
            Setting("AzureApiKey", False) = AzureApiKey
            Setting("AzureRegion", False) = AzureRegion
            Setting("AzureLastLanguagesRetrieve", False) = AzureLastLanguagesRetrieve.ToString("yyyy-MM-dd HH:mm:ss")
            Setting("GoogleApiKey", False) = GoogleApiKey
            Setting("DefaultTargetLocale", False) = DefaultTargetLocale
   Setting("OwnerName", False) = OwnerName
   Setting("OwnerEmail", False) = OwnerEmail
   Setting("OwnerUrl", False) = OwnerUrl
   Setting("OwnerOrganization", False) = OwnerOrganization
   Setting("License", False) = License
   Setting("Copyright", False) = Copyright
   Setting("DefaultDictionary", False) = DefaultDictionary
   Setting("LastUsedDir", False) = LastUsedDir

   Dim isoStore As IsolatedStorageFile = IsolatedStorageFile.GetStore(IsolatedStorageScope.User Or IsolatedStorageScope.Assembly Or IsolatedStorageScope.Domain, Nothing, Nothing)
   Using strOut As New IO.StreamWriter(New IsolatedStorageFileStream(SettingsFilename, FileMode.OpenOrCreate, isoStore))
    Me.WriteXml(strOut, System.Data.XmlWriteMode.IgnoreSchema)
   End Using
  End Sub

 End Class
End Namespace
