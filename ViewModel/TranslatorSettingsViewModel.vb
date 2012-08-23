Namespace ViewModel
 Public Class TranslatorSettingsViewModel
  Inherits ViewModelBase

  Private _appSettings As Common.TranslatorSettings

#Region " Constructors "
  Public Sub New(appSettings As Common.TranslatorSettings)
   _appSettings = appSettings
   Me.DisplayName = "General"
   Me.ID = "General"
  End Sub
#End Region

#Region " Setting Properties "
  Public Property BingClientID() As String
   Get
    Return _appSettings.BingClientID
   End Get
   Set(ByVal value As String)
    _appSettings.BingClientID = value
    Me.OnPropertyChanged("BingClientID")
   End Set
  End Property

  Public Property BingClientSecret() As String
   Get
    Return _appSettings.BingClientSecret
   End Get
   Set(ByVal value As String)
    _appSettings.BingClientSecret = value
    Me.OnPropertyChanged("BingClientSecret")
   End Set
  End Property

  Public Property DefaultTargetLocale() As CultureInfo
   Get
    Return AvailableLocales.FirstOrDefault(Function(x) x.Name = _appSettings.DefaultTargetLocale)
   End Get
   Set(ByVal value As CultureInfo)
    _appSettings.DefaultTargetLocale = value.Name
    Me.OnPropertyChanged("DefaultTargetLocale")
   End Set
  End Property

  Public Property OwnerName() As String
   Get
    Return _appSettings.OwnerName
   End Get
   Set(ByVal value As String)
    _appSettings.OwnerName = value
    Me.OnPropertyChanged("OwnerName")
   End Set
  End Property

  Public Property OwnerEmail() As String
   Get
    Return _appSettings.OwnerEmail
   End Get
   Set(ByVal value As String)
    _appSettings.OwnerEmail = value
    Me.OnPropertyChanged("OwnerEmail")
   End Set
  End Property

  Public Property OwnerUrl() As String
   Get
    Return _appSettings.OwnerUrl
   End Get
   Set(ByVal value As String)
    _appSettings.OwnerUrl = value
    Me.OnPropertyChanged("OwnerUrl")
   End Set
  End Property

  Public Property OwnerOrganization() As String
   Get
    Return _appSettings.OwnerOrganization
   End Get
   Set(ByVal value As String)
    _appSettings.OwnerOrganization = value
    Me.OnPropertyChanged("OwnerOrganization")
   End Set
  End Property

  Public Property License() As String
   Get
    Return _appSettings.License
   End Get
   Set(ByVal value As String)
    _appSettings.License = value
    Me.OnPropertyChanged("License")
   End Set
  End Property

  Public Property Copyright() As String
   Get
    Return _appSettings.Copyright
   End Get
   Set(ByVal value As String)
    _appSettings.Copyright = value
    Me.OnPropertyChanged("Copyright")
   End Set
  End Property
#End Region

#Region " Other Properties "
  Private _availableLocales As List(Of CultureInfo)
  Public ReadOnly Property AvailableLocales() As List(Of CultureInfo)
   Get
    If _availableLocales Is Nothing Then
     _availableLocales = Common.Globals.AllLocales
    End If
    Return _availableLocales
   End Get
  End Property
#End Region

 End Class
End Namespace