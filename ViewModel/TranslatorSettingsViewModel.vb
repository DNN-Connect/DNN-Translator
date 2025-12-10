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
  Private _TranslationProviders As Dictionary(Of String, String)
  Public ReadOnly Property TranslationProviders As Dictionary(Of String, String)
   Get
    If _TranslationProviders Is Nothing Then
     _TranslationProviders = New Dictionary(Of String, String)
     _TranslationProviders.Add("none", "None")
                    _TranslationProviders.Add("bing", "Bing (Microsoft)")
                    _TranslationProviders.Add("google", "Google")
                    _TranslationProviders.Add("azure", "Azure AI Cognitive Services (Microsoft Translator v3)")
                End If
    Return _TranslationProviders
   End Get
  End Property

  Public Property TranslationProvider() As KeyValuePair(Of String, String)
   Get
    For Each kvp As KeyValuePair(Of String, String) In TranslationProviders
     If kvp.Key = _appSettings.TranslationProvider Then Return kvp
    Next
    Return Nothing
   End Get
   Set(ByVal value As KeyValuePair(Of String, String))
    _appSettings.TranslationProvider = value.Key
    Me.OnPropertyChanged("TranslationProvider")
   End Set
  End Property

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

        Public Property GoogleApiKey() As String
            Get
                Return _appSettings.GoogleApiKey
            End Get
            Set(ByVal value As String)
                _appSettings.GoogleApiKey = value
                Me.OnPropertyChanged("GoogleApiKey")
            End Set
        End Property
        Public Property AzureApiKey() As String
            Get
                Return _appSettings.AzureApiKey
            End Get
            Set(ByVal value As String)
                _appSettings.AzureApiKey = value
                Me.OnPropertyChanged("AzureApiKey")
            End Set
        End Property
        Public Property AzureRegion() As String
            Get
                Return _appSettings.AzureRegion
            End Get
            Set(ByVal value As String)
                _appSettings.AzureRegion = value
                Me.OnPropertyChanged("AzureRegion")
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

  Public Property DefaultDictionary() As String
   Get
    Return _appSettings.DefaultDictionary
   End Get
   Set(ByVal value As String)
    _appSettings.DefaultDictionary = value
    Me.OnPropertyChanged("DefaultDictionary")
   End Set
  End Property
#End Region

#Region " Dictionary "
  Private _browseCommand As RelayCommand
  Public ReadOnly Property BrowseCommand As RelayCommand
   Get
    If _browseCommand Is Nothing Then
     _browseCommand = New RelayCommand(Sub(param) Me.BrowseClicked())
    End If
    Return _browseCommand
   End Get
  End Property

  Protected Sub BrowseClicked()

   Dim fbd As New System.Windows.Forms.OpenFileDialog With {.CheckFileExists = True, .DefaultExt = ".dic", .Filter = "dictionary files (*.dic)|*.dic"}
   If _appSettings.LastUsedDir = "" Then
    fbd.InitialDirectory = Common.Globals.TranslatorDocFolder
   Else
    fbd.InitialDirectory = _appSettings.LastUsedDir
   End If
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    If IO.File.Exists(fbd.FileName) Then
     Me.DefaultDictionary = fbd.FileName
    End If
   End If

  End Sub

  Private _newCommand As RelayCommand
  Public ReadOnly Property NewCommand As RelayCommand
   Get
    If _newCommand Is Nothing Then
     _newCommand = New RelayCommand(Sub(param) Me.NewClicked())
    End If
    Return _newCommand
   End Get
  End Property

  Protected Sub NewClicked()

   Dim fbd As New System.Windows.Forms.SaveFileDialog With {.DefaultExt = ".dic", .Filter = "dictionary files (*.dic)|*.dic", .OverwritePrompt = True, .AddExtension = True, .CheckPathExists = True}
   If _appSettings.LastUsedDir = "" Then
    fbd.InitialDirectory = Common.Globals.TranslatorDocFolder
   Else
    fbd.InitialDirectory = _appSettings.LastUsedDir
   End If
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    Me.DefaultDictionary = fbd.FileName
   End If

  End Sub
#End Region

#Region " Other Properties "
  Private _availableLocales As List(Of CultureInfo)
  Public ReadOnly Property AvailableLocales() As List(Of CultureInfo)
   Get
    If _availableLocales Is Nothing Then
     _availableLocales = Common.Globals.AllLocales(CultureTypes.SpecificCultures)
    End If
    Return _availableLocales
   End Get
  End Property
#End Region

 End Class
End Namespace