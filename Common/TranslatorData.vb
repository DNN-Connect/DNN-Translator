Imports System.Text.RegularExpressions

Partial Class TranslatorData

 Public Property SettingsFileName As String = ""

 Public Overridable Sub Save()
  Try
   If Not String.IsNullOrEmpty(SettingsFileName) Then
    Me.WriteXml(SettingsFileName, System.Data.XmlWriteMode.IgnoreSchema)
   End If
  Catch ex As Exception
  End Try
 End Sub

 'Public Sub RefreshResourceFiles()
 ' Me.ResourceFiles.Rows.Clear()
 ' ReadResourceFiles(Location)
 'End Sub

 'Private Sub ReadResourceFiles(directory As String)
 ' For Each f As String In IO.Directory.GetFiles(directory, "*.resx")
 '  Dim m As Match = Regex.Match(f, "\.(\w\w\-\w\w)\.resx")
 '  If m.Success AndAlso m.Groups(1).Value.ToLower <> "en-us" Then
 '  Else
 '   AddResourceFile(f.Replace(Location, ""))
 '  End If
 ' Next
 ' For Each d As String In IO.Directory.GetDirectories(directory)
 '  ReadResourceFiles(d)
 ' Next
 'End Sub

 'Public Sub AddResourceFile(resourcefile As String)
 ' Dim dr As ResourceFilesRow = ResourceFiles.NewResourceFilesRow
 ' dr.FilePath = resourcefile
 ' ResourceFiles.AddResourceFilesRow(dr)
 'End Sub

 Default Public Property Setting(ByVal Key As String, Optional ByVal Save As Boolean = True) As String
  Get
   Dim x As IEnumerable(Of SettingsRow) = From s In Settings
                     Select s Where s.SettingKey = Key
   If x.Count = 0 Then
    Return Nothing
   Else
    Return x(0).SettingValue
   End If
  End Get
  Set(ByVal Value As String)
   Dim x As IEnumerable(Of SettingsRow) = From s In Settings
                     Select s Where s.SettingKey = Key
   If x.Count = 0 Then
    Dim s As SettingsRow = Settings.NewSettingsRow
    s.SettingKey = Key
    s.SettingValue = Value
    Settings.AddSettingsRow(s)
   Else
    x(0).SettingValue = Value
   End If
   If Save Then
    Me.Save()
   End If
  End Set
 End Property

 Public Sub ReadSettingValue(name As String, ByRef var As String)
  Dim res As String = Setting(name)
  If Not String.IsNullOrEmpty(res) Then var = res
 End Sub

 Public Sub ReadSettingValue(name As String, ByRef var As Boolean)
  Dim res As String = Setting(name)
  If Not String.IsNullOrEmpty(res) Then var = Boolean.Parse(res)
 End Sub

 Public Sub ReadSettingValue(name As String, ByRef var As Date)
  Dim res As String = Setting(name)
  If Not String.IsNullOrEmpty(res) Then var = Date.Parse(res)
 End Sub

End Class
