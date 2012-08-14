Imports System.IO.IsolatedStorage
Imports System.IO

Namespace Common
 Public Class TranslatorSettings
  Inherits TranslatorData

  Private Shadows Const SettingsFilename As String = "TranslatorSettings.xml"

  Public Property BingClientID As String = ""
  Public Property BingClientSecret As String = ""
  Public Property DefaultTargetLocale As String = ""
  Public Property BingLocales As New List(Of CultureInfo)
  Public Property MaximumKeys As Integer = 100

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

   For Each l As BingLanguagesRow In BingLanguages.Rows
    BingLocales.Add(New CultureInfo(l.Locale))
   Next

   ReadSettingValue("BingClientID", BingClientID)
   ReadSettingValue("BingClientSecret", BingClientSecret)
   ReadSettingValue("DefaultTargetLocale", DefaultTargetLocale)

  End Sub

  Public Overrides Sub Save()

   BingLanguages.Rows.Clear()
   For Each l As CultureInfo In BingLocales
    Dim lr As BingLanguagesRow = BingLanguages.NewBingLanguagesRow
    lr.Locale = l.Name
    BingLanguages.AddBingLanguagesRow(lr)
   Next

   Setting("BingClientID", False) = BingClientID
   Setting("BingClientSecret", False) = BingClientSecret
   Setting("DefaultTargetLocale", False) = DefaultTargetLocale

   Dim isoStore As IsolatedStorageFile = IsolatedStorageFile.GetStore(IsolatedStorageScope.User Or IsolatedStorageScope.Assembly Or IsolatedStorageScope.Domain, Nothing, Nothing)
   Using strOut As New IO.StreamWriter(New IsolatedStorageFileStream(SettingsFilename, FileMode.OpenOrCreate, isoStore))
    Me.WriteXml(strOut, System.Data.XmlWriteMode.IgnoreSchema)
   End Using
  End Sub

 End Class
End Namespace
