Imports System.Collections.ObjectModel

Namespace ViewModel
 Public Class SettingsViewModel
  Inherits ViewModelBase

  Private _parentViewModel As MainWindowViewModel
  Friend ReadOnly Property ParentViewModel As MainWindowViewModel
   Get
    Return _parentViewModel
   End Get
  End Property

  Private _translatorSettings As Common.TranslatorSettings
  Property TranslatorSettings As Common.TranslatorSettings
   Get
    Return _translatorSettings
   End Get
   Set(value As Common.TranslatorSettings)
    _translatorSettings = value
    Me.OnPropertyChanged("TranslatorSettings")
   End Set
  End Property

  Private _projectSettings As Common.ProjectSettings
  Public Property ProjectSettings() As Common.ProjectSettings
   Get
    Return _projectSettings
   End Get
   Set(ByVal value As Common.ProjectSettings)
    _projectSettings = value
    Me.OnPropertyChanged("ProjectSettings")
   End Set
  End Property


#Region " Commands "
  Private _cancelCommand As RelayCommand
  Public ReadOnly Property CancelCommand As RelayCommand
   Get
    If _cancelCommand Is Nothing Then
     _cancelCommand = New RelayCommand(Sub(param) Me.OnRequestCancel(param))
    End If
    Return _cancelCommand
   End Get
  End Property
  Public Event RequestCancel As EventHandler
  Protected Sub OnRequestCancel(o As Object)
   RaiseEvent RequestCancel(Me, EventArgs.Empty)
   CType(o, Window).Close()
  End Sub

  Private _okCommand As RelayCommand
  Public ReadOnly Property OkCommand As RelayCommand
   Get
    If _okCommand Is Nothing Then
     _okCommand = New RelayCommand(Sub(param) Me.OnRequestOk(param))
    End If
    Return _okCommand
   End Get
  End Property
  Public Event RequestOk As EventHandler
  Protected Sub OnRequestOk(o As Object)
   ' saving back the settings to the parent and to their respective files
   If _parentViewModel.TranslatorSettings.BingClientID <> TranslatorSettings.BingClientID Then
    ' we changed the Bing Client ID
    _parentViewModel.Bing = New Common.Bing.Bing(TranslatorSettings)
    If TranslatorSettings.BingClientID <> "" Then
     Try
      TranslatorSettings.BingLocales = _parentViewModel.Bing.GetLanguages
     Catch ex As Exception
     End Try
    End If
   End If
   _parentViewModel.TranslatorSettings = TranslatorSettings
   TranslatorSettings.Save()
   If ProjectSettings IsNot Nothing Then
    _parentViewModel.ProjectSettings = ProjectSettings
    ProjectSettings.Save()
   End If
   RaiseEvent RequestOk(Me, EventArgs.Empty)
   CType(o, Window).Close()
  End Sub
#End Region

#Region " Panels "
  Public Property SettingPanels As New ObservableCollection(Of ViewModelBase)
  Private _selectedPanel As Object = Nothing
  Public Property SelectedPanel As Object
   Get
    Return _selectedPanel
   End Get
   Set(value As Object)
    _selectedPanel = value
    Me.OnPropertyChanged("SelectedPanel")
   End Set
  End Property

  Private _selectTabCommand As RelayCommand
  Public ReadOnly Property SelectTabCommand As RelayCommand
   Get
    If _selectTabCommand Is Nothing Then
     _selectTabCommand = New RelayCommand(Sub(param) SelectTabCommandHandler(param))
    End If
    Return _selectTabCommand
   End Get
  End Property
  Protected Sub SelectTabCommandHandler(param As Object)
   If TypeOf param Is ViewModelBase Then
    SelectedPanel = param
   End If
  End Sub
#End Region

#Region " Constructors "
  Public Sub New(mainWindowViewModel As MainWindowViewModel)
   _parentViewModel = mainWindowViewModel
   _translatorSettings = CType(mainWindowViewModel.TranslatorSettings.Clone, Common.TranslatorSettings)
   _projectSettings = mainWindowViewModel.ProjectSettings
   SettingPanels.Add(New TranslatorSettingsViewModel(Me))
   If _projectSettings IsNot Nothing Then
    SettingPanels.Add(New ProjectSettingsViewModel(Me))
   End If
   _selectedPanel = SettingPanels(0)
  End Sub
#End Region


 End Class
End Namespace
