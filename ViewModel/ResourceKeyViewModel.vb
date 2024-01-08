Imports DotNetNuke.Translator.Common
Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq

Namespace ViewModel
  Public Class ResourceKeyViewModel
    Inherits ViewModelBase

#Region " Properties "
    Private _resourceFile As String
    Public ReadOnly Property ResourceFile() As String
      Get
        Return _resourceFile
      End Get
    End Property

    Private _key As String
    Public ReadOnly Property Key As String
      Get
        Return _key
      End Get
    End Property

    Private _originalValue As String
    Public ReadOnly Property OriginalValue As String
      Get
        Return _originalValue
      End Get
    End Property

    Private _targetValue As String
    Public Property TargetValue() As String
      Get
        Return _targetValue
      End Get
      Set(ByVal value As String)
        _targetValue = value
        _changed = True
        Downloaded = False

        Me.SetTargetColor()
        Me.OnPropertyChanged("TargetValue")
        Me.OnPropertyChanged("Changed")
        Me.OnPropertyChanged("TargetColor")
        If _LEText IsNot Nothing Then
          _LEText.Translation = Microsoft.Security.Application.UnicodeCharacterEncoder.XmlEncode(value)
        End If
      End Set
    End Property

    Private _compareValue As String
    Public Property CompareValue() As String
      Get
        Return _compareValue
      End Get
      Set(ByVal value As String)
        _compareValue = value
        Me.OnPropertyChanged("CompareValue")
      End Set
    End Property

    Private _selected As Boolean = False
    Public Property Selected() As Boolean
      Get
        Return _selected
      End Get
      Set(ByVal value As Boolean)
        _selected = value
        Me.OnPropertyChanged("Selected")
      End Set
    End Property

    Private _TargetFocused As Boolean = False
    Public Property TargetFocused() As Boolean
      Get
        Return _TargetFocused
      End Get
      Set(ByVal value As Boolean)
        _TargetFocused = value
        Me.OnPropertyChanged("TargetFocused")
      End Set
    End Property

    Private _comment As String
    Public Property Comment() As String
      Get
        Return _comment
      End Get
      Set(ByVal value As String)
        _comment = value
        Me.OnPropertyChanged("Comment")
      End Set
    End Property

    Private _changed As Boolean
    Public Property Changed As Boolean
      Get
        Return _changed
      End Get
      Set(value As Boolean)
        _changed = value
      End Set
    End Property

    Private _TranslatedResource As Resource
    Public Property TranslatedResource() As Resource
      Get
        Return _TranslatedResource
      End Get
      Set(ByVal value As Resource)
        _TranslatedResource = value
        _targetValue = _TranslatedResource.Value
        _LastModified = _TranslatedResource.LastModified
      End Set
    End Property

    Public Property LEText As Common.LEService.TextInfo
    Public Property OriginalResource() As Resource
    Public Property Downloaded As Boolean = False

    Private _targetColor As SolidColorBrush
    Public Property TargetColor() As SolidColorBrush
      Get
        If _targetColor Is Nothing Then
          Me.SetTargetColor()
        End If
        Return _targetColor
      End Get
      Set(ByVal value As SolidColorBrush)
        _targetColor = value
        Me.OnPropertyChanged("TargetColor")
      End Set
    End Property

    Private _compareColor As SolidColorBrush
    Public Property CompareColor() As SolidColorBrush
      Get
        'If _compareColor Is Nothing Then _compareColor = Brushes.White
        If _compareColor Is Nothing Then _compareColor = CType((New BrushConverter).ConvertFrom(_compareC), SolidColorBrush)
        Return _compareColor
      End Get
      Set(ByVal value As SolidColorBrush)
        _compareColor = value
        Me.OnPropertyChanged("CompareColor")
      End Set
    End Property

    Public Property LastModified As DateTime = DateTime.MinValue
#End Region

#Region " Constructors "
    Public Sub New(translatedResource As Resource, leText As Common.LEService.TextInfo)
      _TranslatedResource = translatedResource
      _LEText = leText
      _resourceFile = leText.FilePath
      _key = leText.TextKey
      _originalValue = Web.HttpUtility.HtmlDecode(leText.OriginalValue)
      If translatedResource IsNot Nothing Then
        _targetValue = translatedResource.Value
      End If
      _compareValue = Web.HttpUtility.HtmlDecode(leText.Translation)
      Me.ID = _resourceFile & ":" & _key
    End Sub

    Public Sub New(originalResource As Resource)
      _OriginalResource = originalResource
      _resourceFile = originalResource.FileKey
      _key = originalResource.Key
      _originalValue = originalResource.Value
      Me.ID = _resourceFile & ":" & _key
    End Sub

    Public Sub New(originalResource As Resource, translation As String)
      _OriginalResource = originalResource
      _resourceFile = originalResource.FileKey
      _key = originalResource.Key
      _originalValue = originalResource.Value
      _targetValue = translation
      Me.ID = _resourceFile & ":" & _key
    End Sub

    Public Sub New(originalResource As Resource, translatedResource As Resource)
      _OriginalResource = originalResource
      _TranslatedResource = translatedResource
      _resourceFile = originalResource.FileKey
      _key = originalResource.Key
      _originalValue = originalResource.Value
      If translatedResource IsNot Nothing Then
        _targetValue = translatedResource.Value
      End If
      Me.ID = _resourceFile & ":" & _key
    End Sub

    Public Sub New(originalResource As Resource, translatedResource As Resource, compareValue As String)
      _OriginalResource = originalResource
      _TranslatedResource = translatedResource
      _compareValue = compareValue

      _resourceFile = originalResource.FileKey
      _key = originalResource.Key
      _originalValue = originalResource.Value
      If translatedResource IsNot Nothing Then
        _targetValue = translatedResource.Value
      End If
      Me.ID = _resourceFile & ":" & _key
    End Sub

    Public Sub New(originalResource As Resource, translation As String, compareValue As String)
      _OriginalResource = originalResource
      _compareValue = compareValue
      _resourceFile = originalResource.FileKey
      _key = originalResource.Key
      _originalValue = originalResource.Value
      _targetValue = translation
      Me.ID = _resourceFile & ":" & _key
    End Sub

    Public Sub New(originalResource As Resource, translatedResource As Resource, compareValue As String, leText As Common.LEService.TextInfo)
      _OriginalResource = originalResource
      _TranslatedResource = translatedResource
      _compareValue = compareValue
      _LEText = leText

      _resourceFile = originalResource.FileKey
      _key = originalResource.Key
      _originalValue = originalResource.Value
      If translatedResource IsNot Nothing Then
        _targetValue = translatedResource.Value
      End If
      Me.ID = _resourceFile & ":" & _key
    End Sub
#End Region

    Private _targetC As String = "#FFFFFF"
    Public Sub HighlightTarget()
      HighlightTarget("#F2DEDE")
    End Sub
    Public Sub HighlightTarget(color As String)
      _targetC = color
    End Sub

    Private _compareC As String = "#FFFFFF"
    Public Sub HighlightCompare()
      HighlightCompare("#F2DEDE")
    End Sub
    Public Sub HighlightCompare(color As String)
      _compareC = color
    End Sub

    Public Function Clone() As ResourceKeyViewModel
      Dim res As New ResourceKeyViewModel(OriginalResource, TranslatedResource, CompareValue, LEText)
      res.TargetValue = TargetValue
      Return res
    End Function

    Public Sub SetDownloadedValue(text As Common.LEService.TextInfo)
      _targetValue = Web.HttpUtility.HtmlDecode(text.Translation)
      _LastModified = text.LastModified
      Me.OnPropertyChanged("TargetValue")
      Me.OnPropertyChanged("TargetColor")
    End Sub

    Private Sub SetTargetColor()
      Dim col As String = ""
      If _targetValue = "" Then
        col = "#faed5f"
      Else
        If _targetValue = OriginalValue Then
          col = "#F2DEDE"
        End If
        If _targetValue.IndexOf("<") >= 0 Then
          Dim d = New HtmlDocument
          d.LoadHtml(_targetValue)
          If d.ParseErrors.Count > 0 Then
            col = "#f77b45"
          End If
        End If
      End If
      If col = "" Then
        _targetColor = CType((New BrushConverter).ConvertFrom(_targetC), SolidColorBrush)
      Else
        _targetColor = CType((New BrushConverter).ConvertFrom(col), SolidColorBrush)
      End If
    End Sub

  End Class
End Namespace
