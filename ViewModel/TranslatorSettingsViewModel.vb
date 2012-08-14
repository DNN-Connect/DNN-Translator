Namespace ViewModel
 Public Class TranslatorSettingsViewModel
  Inherits ViewModelBase

  Private _settingsViewModel As SettingsViewModel

  Public Sub New(settingsViewModel As SettingsViewModel)
   _settingsViewModel = settingsViewModel
   Me.DisplayName = "General"
   Me.ID = "General"
  End Sub

  Public Property BingClientID() As String
   Get
    Return _settingsViewModel.TranslatorSettings.BingClientID
   End Get
   Set(ByVal value As String)
    _settingsViewModel.TranslatorSettings.BingClientID = value
    Me.OnPropertyChanged("BingClientID")
   End Set
  End Property

  Public Property BingClientSecret() As String
   Get
    Return _settingsViewModel.TranslatorSettings.BingClientSecret
   End Get
   Set(ByVal value As String)
    _settingsViewModel.TranslatorSettings.BingClientSecret = value
    Me.OnPropertyChanged("BingClientSecret")
   End Set
  End Property

  Public Property DefaultTargetLocale() As CultureInfo
   Get
    Return AvailableLocales.FirstOrDefault(Function(x) x.Name = _settingsViewModel.TranslatorSettings.DefaultTargetLocale)
   End Get
   Set(ByVal value As CultureInfo)
    _settingsViewModel.TranslatorSettings.DefaultTargetLocale = value.Name
    Me.OnPropertyChanged("DefaultTargetLocale")
   End Set
  End Property

  Private _availableLocales As List(Of CultureInfo)
  Public ReadOnly Property AvailableLocales() As List(Of CultureInfo)
   Get
    If _availableLocales Is Nothing Then
     _availableLocales = Common.Globals.AllLocales
    End If
    Return _availableLocales
   End Get
  End Property

 End Class
End Namespace