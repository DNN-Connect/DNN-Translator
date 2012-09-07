Imports System.Data.SqlClient
Imports DotNetNuke.Translator.ViewModel

Namespace Common
 Public Class ProjectSettings
  Inherits TranslatorData

  ' User configurable
  Public Property Location As String = ""
  Public Property TargetLocale As String = ""
  Public Property MappedLocale As String = ""
  Public Property ConnectionUrl As String = ""
  Public Property Username As String = ""
  Public Property Password As Security.SecureString
  Public Property LocalUrl As String = ""
  Public Property OverrideOwner As Boolean = False
  Public Property OwnerName As String = ""
  Public Property OwnerEmail As String = ""
  Public Property OwnerUrl As String = ""
  Public Property OwnerOrganization As String = ""
  Public Property License As String = ""
  Public Property Copyright As String = ""
  Public Property Dictionary As String = ""

  ' Parsed on load
  Public Property DotNetNukeVersion As String = ""
  Public Property DotNetNukeType As String = "Community Edition"
  Public Property WebConfig As New Xml.XmlDocument
  Public Property ConnectionString As String = ""
  Public Property InstalledPackages As New List(Of InstalledPackageViewModel)
  Public Property InstalledLanguages As New List(Of String)
  Public Property AvailableLocales As New List(Of CultureInfo)
  Public Property CurrentSnapShot As Snapshot
  Private Property DatabaseOwner As String = ""
  Private Property ObjectQualifier As String = ""

  Public Sub New(projectSettingsFileOrDnnDirectory As String)
   MyBase.New()

   If IO.File.Exists(projectSettingsFileOrDnnDirectory) Then ' it's a settings file
    SettingsFileName = projectSettingsFileOrDnnDirectory
    Try
     Me.ReadXml(SettingsFileName, System.Data.XmlReadMode.IgnoreSchema)
    Catch ex As Exception
    End Try
    ReadSettingValue("Location", Location)
    ReadSettingValue("TargetLocale", TargetLocale)
    ReadSettingValue("MappedLocale", MappedLocale)
    If Setting("AvailableLocales") IsNot Nothing Then
     For Each l As String In Setting("AvailableLocales").Split(";"c)
      AvailableLocales.Add(New CultureInfo(l))
     Next
    End If
    ReadSettingValue("ConnectionUrl", ConnectionUrl)
    ReadSettingValue("Username", Username)
    If Setting("Password") IsNot Nothing Then
     Try
      Password = Common.Globals.DecryptString(Setting("Password"))
     Catch ex As Exception
     End Try
    End If
    ReadSettingValue("LocalUrl", LocalUrl)
    ReadSettingValue("OverrideOwner", OverrideOwner)
    ReadSettingValue("OwnerName", OwnerName)
    ReadSettingValue("OwnerEmail", OwnerEmail)
    ReadSettingValue("OwnerUrl", OwnerUrl)
    ReadSettingValue("OwnerOrganization", OwnerOrganization)
    ReadSettingValue("License", License)
    ReadSettingValue("Copyright", Copyright)
    ReadSettingValue("Dictionary", Dictionary)
   ElseIf IO.Directory.Exists(projectSettingsFileOrDnnDirectory) Then ' it's a DNN directory
    If Not projectSettingsFileOrDnnDirectory.EndsWith("\") Then projectSettingsFileOrDnnDirectory &= "\"
    Me.Location = projectSettingsFileOrDnnDirectory
   End If

   CurrentSnapShot = New Snapshot(Location, Location)

   Try
    ' Load various DNN and reflect
    Dim ass As System.Reflection.Assembly = System.Reflection.Assembly.LoadFile(Location & "bin\dotnetnuke.dll")
    DotNetNukeVersion = ass.GetName.Version.ToString
    If IO.File.Exists(Location & "bin\DotNetNuke.Professional.dll") Then
     DotNetNukeType = "Professional Edition"
    End If
    If IO.File.Exists(Location & "bin\DotNetNuke.Enterprise.dll") Then
     DotNetNukeType = "Enterprise Edition"
    End If

    Dim core As New InstalledPackageViewModel("Core", DotNetNukeType, DotNetNukeVersion, "")
    InstalledPackages.Add(core)

    ' Load web.config
    WebConfig.Load(Location & "web.config")
    ConnectionString = WebConfig.SelectSingleNode("/configuration/connectionStrings/add[@name='SiteSqlServer']/@connectionString").InnerText
    DatabaseOwner = WebConfig.SelectSingleNode("/configuration/dotnetnuke/data/providers/add[@name='SqlDataProvider']/@databaseOwner").InnerText
    If DatabaseOwner <> "" AndAlso Not DatabaseOwner.EndsWith(".") Then DatabaseOwner &= "."
    ObjectQualifier = WebConfig.SelectSingleNode("/configuration/dotnetnuke/data/providers/add[@name='SqlDataProvider']/@objectQualifier").InnerText
    If ObjectQualifier <> "" AndAlso Not ObjectQualifier.EndsWith("_") Then ObjectQualifier &= "_"

    ' Try to get a list of installed packages and attempt to build the resource file list of each object
    Using c As New SqlConnection(ConnectionString)
     c.Open()
     Dim q As SqlCommand = c.CreateCommand
     q.CommandText = String.Format("SELECT *, dm.FolderName ModuleFolderName FROM {0}{1}Packages p LEFT JOIN {0}{1}DesktopModules dm ON dm.PackageID=p.PackageID", DatabaseOwner, ObjectQualifier)
     Using ir As SqlDataReader = q.ExecuteReader
      Do While ir.Read
       Dim package As New InstalledPackageViewModel(ir)
       If package.LoadManifest(Me.Location) Then
        InstalledPackages.Add(package)
       End If
      Loop
     End Using
     q = c.CreateCommand
     q.CommandText = String.Format("SELECT * FROM {0}{1}Languages", DatabaseOwner, ObjectQualifier)
     Using ir As SqlDataReader = q.ExecuteReader
      Do While ir.Read
       InstalledLanguages.Add(CStr(ir.Item("CultureCode")))
      Loop
     End Using
    End Using

    ' Now adjust the list for the core
    Dim allNonCoreResourceFiles As New List(Of String)
    For Each p As InstalledPackageViewModel In InstalledPackages
     If p.Manifest IsNot Nothing Then
      allNonCoreResourceFiles.AddRange(p.Manifest.ResourceFiles)
     End If
    Next
    core.Manifest = New ModuleManifest
    For Each resFile As String In CurrentSnapShot.ResourceFiles.Keys
     If Not allNonCoreResourceFiles.Contains(resFile) Then
      core.Manifest.ResourceFiles.Add(resFile)
     End If
    Next

   Catch ex As Exception

   End Try

  End Sub

  Public Overloads Sub Save(projectSettingsFile As String)

   If projectSettingsFile <> "" Then SettingsFileName = projectSettingsFile

   Setting("TargetLocale", False) = TargetLocale
   Setting("MappedLocale", False) = MappedLocale
   Dim al As New List(Of String)
   For Each ci As CultureInfo In AvailableLocales
    al.Add(ci.Name)
   Next
   Setting("AvailableLocales", False) = String.Join(";", al)
   Setting("ConnectionUrl", False) = ConnectionUrl
   Setting("Username", False) = Username
   Setting("Password", False) = Common.Globals.EncryptString(Password)
   Setting("LocalUrl", False) = LocalUrl
   Setting("OverrideOwner", False) = OverrideOwner.ToString
   Setting("OwnerName", False) = OwnerName
   Setting("OwnerEmail", False) = OwnerEmail
   Setting("OwnerUrl", False) = OwnerUrl
   Setting("OwnerOrganization", False) = OwnerOrganization
   Setting("License", False) = License
   Setting("Copyright", False) = Copyright
   Setting("Dictionary", False) = Dictionary

   MyBase.Save()

  End Sub

  Public Overrides Sub Save()
   Save("")
  End Sub

 End Class
End Namespace