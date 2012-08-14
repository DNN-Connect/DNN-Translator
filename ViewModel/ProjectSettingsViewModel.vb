Namespace ViewModel
 Public Class ProjectSettingsViewModel
  Inherits ViewModelBase

  Private _settingsViewModel As SettingsViewModel

  Public Sub New(settingsViewModel As SettingsViewModel)
   _settingsViewModel = settingsViewModel
   Me.DisplayName = "Project"
   Me.ID = "Project"
  End Sub

  Public Property TargetLocale() As CultureInfo
   Get
    Return AllLocales.FirstOrDefault(Function(x) x.Name = _settingsViewModel.ProjectSettings.TargetLocale)
   End Get
   Set(ByVal value As CultureInfo)
    _settingsViewModel.ProjectSettings.TargetLocale = value.Name
    Me.OnPropertyChanged("TargetLocale")
    If value.Name.Length > 2 Then
     MappedLocale = value
    End If
   End Set
  End Property

  Public Property MappedLocale() As CultureInfo
   Get
    Return AvailableLocales.FirstOrDefault(Function(x) x.Name = _settingsViewModel.ProjectSettings.MappedLocale)
   End Get
   Set(ByVal value As CultureInfo)
    _settingsViewModel.ProjectSettings.MappedLocale = value.Name
    Me.OnPropertyChanged("MappedLocale")
   End Set
  End Property

  Private _allLocales As List(Of CultureInfo)
  Public ReadOnly Property AllLocales() As List(Of CultureInfo)
   Get
    If _allLocales Is Nothing Then
     _allLocales = Common.Globals.AllLocales(CultureTypes.SpecificCultures)
    End If
    Return _allLocales
   End Get
  End Property

  Public Property AvailableLocales() As List(Of CultureInfo)
   Get
    Return _settingsViewModel.ProjectSettings.AvailableLocales
   End Get
   Set(value As List(Of CultureInfo))
    _settingsViewModel.ProjectSettings.AvailableLocales = value
   End Set
  End Property

  Public Property ConnectionUrl() As String
   Get
    Return _settingsViewModel.ProjectSettings.ConnectionUrl
   End Get
   Set(ByVal value As String)
    _settingsViewModel.ProjectSettings.ConnectionUrl = value
   End Set
  End Property

  Public Property Username() As String
   Get
    Return _settingsViewModel.ProjectSettings.Username
   End Get
   Set(ByVal value As String)
    _settingsViewModel.ProjectSettings.Username = value
   End Set
  End Property

  Public Property Password() As String
   Get
    Return Common.Globals.ToInsecureString(_settingsViewModel.ProjectSettings.Password)
   End Get
   Set(ByVal value As String)
    _settingsViewModel.ProjectSettings.Password = Common.Globals.ToSecureString(value)
   End Set
  End Property

  Public Property LocalUrl() As String
   Get
    Return _settingsViewModel.ProjectSettings.LocalUrl
   End Get
   Set(ByVal value As String)
    _settingsViewModel.ProjectSettings.LocalUrl = value
   End Set
  End Property

#Region " Testing "
  Private _TestConnection As RelayCommand
  Public ReadOnly Property TestConnectionCommand As RelayCommand
   Get
    If _TestConnection Is Nothing Then
     _TestConnection = New RelayCommand(Sub(param) Me.TestConnectionClicked(param), Function(param) Me.TestConnectionEnabled)
    End If
    Return _TestConnection
   End Get
  End Property

  Protected Function TestConnectionEnabled() As Boolean
   Return CBool(ConnectionUrl <> "" And Username <> "" And Password <> "")
  End Function

  Protected Sub TestConnectionClicked(param As Object)

   Dim service As New Common.LEService.LEService(ConnectionUrl, Username, _settingsViewModel.ProjectSettings.Password)
   Try
    Dim locs As List(Of CultureInfo) = service.GetEditLanguages
    AvailableLocales = locs
    Me.OnPropertyChanged("TargetLocale")
    Me.OnPropertyChanged("AvailableLocales")
   Catch ex As Exception
   End Try

  End Sub
#End Region
 End Class
End Namespace