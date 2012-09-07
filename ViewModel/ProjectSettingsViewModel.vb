Namespace ViewModel
 Public Class ProjectSettingsViewModel
  Inherits ViewModelBase

  Private _projectSettings As Common.ProjectSettings

#Region " Constructor "
  Public Sub New(projectSettings As Common.ProjectSettings)
   _projectSettings = projectSettings
   Me.DisplayName = "Project Settings"
   Me.ID = "Project"
  End Sub
#End Region

#Region " Setting Properties "
  Public Property Location() As String
   Get
    Return _projectSettings.Location
   End Get
   Set(ByVal value As String)
    _projectSettings.Location = value
   End Set
  End Property

  Public Property TargetLocale() As CultureInfo
   Get
    Return AllLocales.FirstOrDefault(Function(x) x.Name = _projectSettings.TargetLocale)
   End Get
   Set(ByVal value As CultureInfo)
    _projectSettings.TargetLocale = value.Name
    Me.OnPropertyChanged("TargetLocale")
    If value.Name.Length > 2 Then
     MappedLocale = value
    End If
   End Set
  End Property

  Public Property MappedLocale() As CultureInfo
   Get
    Return AvailableLocales.FirstOrDefault(Function(x) x.Name = _projectSettings.MappedLocale)
   End Get
   Set(ByVal value As CultureInfo)
    _projectSettings.MappedLocale = value.Name
    Me.OnPropertyChanged("MappedLocale")
   End Set
  End Property

  Public Property ConnectionUrl() As String
   Get
    Return _projectSettings.ConnectionUrl
   End Get
   Set(ByVal value As String)
    _projectSettings.ConnectionUrl = value
    Me.OnPropertyChanged("ConnectionUrl")
    If value.Trim = "" Then
     HasRemote = False
    Else
     HasRemote = True
    End If
   End Set
  End Property

  Public Property Username() As String
   Get
    Return _projectSettings.Username
   End Get
   Set(ByVal value As String)
    _projectSettings.Username = value
    Me.OnPropertyChanged("Username")
   End Set
  End Property

  Public Property Password() As String
   Get
    Return Common.Globals.ToInsecureString(_projectSettings.Password)
   End Get
   Set(ByVal value As String)
    _projectSettings.Password = Common.Globals.ToSecureString(value)
    Me.OnPropertyChanged("Password")
   End Set
  End Property

  Public Property LocalUrl() As String
   Get
    Return _projectSettings.LocalUrl
   End Get
   Set(ByVal value As String)
    _projectSettings.LocalUrl = value
    Me.OnPropertyChanged("LocalUrl")
   End Set
  End Property

  Public Property OverrideOwner() As Boolean
   Get
    Return _projectSettings.OverrideOwner
   End Get
   Set(ByVal value As Boolean)
    _projectSettings.OverrideOwner = value
    Me.OnPropertyChanged("OverrideOwner")
   End Set
  End Property

  Public Property OwnerName() As String
   Get
    Return _projectSettings.OwnerName
   End Get
   Set(ByVal value As String)
    _projectSettings.OwnerName = value
    Me.OnPropertyChanged("OwnerName")
   End Set
  End Property

  Public Property OwnerEmail() As String
   Get
    Return _projectSettings.OwnerEmail
   End Get
   Set(ByVal value As String)
    _projectSettings.OwnerEmail = value
    Me.OnPropertyChanged("OwnerEmail")
   End Set
  End Property

  Public Property OwnerUrl() As String
   Get
    Return _projectSettings.OwnerUrl
   End Get
   Set(ByVal value As String)
    _projectSettings.OwnerUrl = value
    Me.OnPropertyChanged("OwnerUrl")
   End Set
  End Property

  Public Property OwnerOrganization() As String
   Get
    Return _projectSettings.OwnerOrganization
   End Get
   Set(ByVal value As String)
    _projectSettings.OwnerOrganization = value
    Me.OnPropertyChanged("OwnerOrganization")
   End Set
  End Property

  Public Property License() As String
   Get
    Return _projectSettings.License
   End Get
   Set(ByVal value As String)
    _projectSettings.License = value
    Me.OnPropertyChanged("License")
   End Set
  End Property

  Public Property Copyright() As String
   Get
    Return _projectSettings.Copyright
   End Get
   Set(ByVal value As String)
    _projectSettings.Copyright = value
    Me.OnPropertyChanged("Copyright")
   End Set
  End Property

  Public Property Dictionary() As String
   Get
    Return _projectSettings.Dictionary
   End Get
   Set(ByVal value As String)
    _projectSettings.Dictionary = value
    Me.OnPropertyChanged("Dictionary")
   End Set
  End Property
#End Region

#Region " Other Properties "
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
    Return _projectSettings.AvailableLocales
   End Get
   Set(value As List(Of CultureInfo))
    _projectSettings.AvailableLocales = value
   End Set
  End Property

  Private _hasRemote As Boolean
  Public Property HasRemote() As Boolean
   Get
    Return _hasRemote
   End Get
   Set(ByVal value As Boolean)
    _hasRemote = value
    Me.OnPropertyChanged("HasRemote")
   End Set
  End Property

#End Region

#Region " Test Remote Connection "
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

   Dim service As New Common.LEService.LEService(ConnectionUrl, Username, _projectSettings.Password)
   Try
    Dim locs As List(Of CultureInfo) = service.GetEditLanguages
    AvailableLocales = locs
    Me.OnPropertyChanged("TargetLocale")
    Me.OnPropertyChanged("AvailableLocales")
   Catch ex As Exception
   End Try

  End Sub
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
   fbd.InitialDirectory = IO.Path.GetDirectoryName(Location)
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    If IO.File.Exists(fbd.FileName) Then
     Me.Dictionary = fbd.FileName
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
   fbd.InitialDirectory = IO.Path.GetDirectoryName(Location)
   Dim result As System.Windows.Forms.DialogResult = fbd.ShowDialog()
   If result = Forms.DialogResult.OK Or result = Forms.DialogResult.Yes Then
    Me.Dictionary = fbd.FileName
   End If

  End Sub
#End Region

 End Class
End Namespace