Imports System.ServiceModel.Channels
Imports System.ServiceModel
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Xml

Namespace Services.Translation.Bing
 Public Class BingTranslationService
  Implements ITranslationService

  Private _bingClientId As String = ""
  Private _bingClientSecret As String = ""
  Private _accessToken As AdmAccessToken = Nothing
  Private _SupportedLanguages As List(Of CultureInfo)

#Region " Contructors "
  Public Sub New(appSettings As Common.TranslatorSettings)
   _bingClientId = appSettings.BingClientID
   _bingClientSecret = appSettings.BingClientSecret

   If appSettings.BingLastLanguagesRetrieve < Now.AddMonths(-1) OrElse appSettings.BingLanguages = "" Then
    Task.Run(Function() Me.RetrieveLanguages).Wait()
    If _SupportedLanguages.Count > 0 Then
     Dim _cache As String = ""
     For Each ci As CultureInfo In _SupportedLanguages
      _cache &= ci.Name & ","
     Next
     appSettings.BingLanguages = _cache.TrimEnd(","c)
     appSettings.BingLastLanguagesRetrieve = Now
     appSettings.Save()
    End If
   Else
    _SupportedLanguages = New List(Of CultureInfo)
    For Each l As String In appSettings.BingLanguages.Split(","c)
     Try
      _SupportedLanguages.Add(New CultureInfo(l))
     Catch ex As Exception
     End Try
    Next
   End If

  End Sub
#End Region

#Region " Private Methods "
  Private ReadOnly Property AccessToken As String
   Get
    If String.IsNullOrEmpty(_bingClientId) Then Return ""
    If _accessToken Is Nothing OrElse _accessToken.Expires < Now Then
     Dim admAuth = New AdmAuthentication(_bingClientId, _bingClientSecret)
     Try
      _accessToken = admAuth.GetAccessToken()
      _accessToken.Expires = DateTime.Now.AddSeconds(Integer.Parse(_accessToken.ExpiresIn))
     Catch e As Exception
      Return ""
     End Try
    End If
    Return "Bearer " & _accessToken.AccessToken
   End Get
  End Property

  Private Async Function RetrieveLanguages() As Task(Of Boolean)
   Dim res As New List(Of CultureInfo)
   Dim _Authentication = New AzureDataMarket(_bingClientId, _bingClientSecret)
   Dim m_Token As AzureDataMarket.Token = Await _Authentication.GetTokenAsync()
   Dim auth As String = m_Token.Header
   Dim client As New HttpClient()
   client.DefaultRequestHeaders.Add("Authorization", auth)
   Dim RequestUri As String = "http://api.microsofttranslator.com/v2/Http.svc/GetLanguagesForTranslate?appId="
   Try
    Dim ret As Byte() = Await client.GetByteArrayAsync(RequestUri)
    Dim dcs As New System.Runtime.Serialization.DataContractSerializer(GetType(List(Of String)))
    Using reader As XmlDictionaryReader = XmlDictionaryReader.CreateTextReader(ret, New XmlDictionaryReaderQuotas())
     Dim lngs As List(Of String) = CType(dcs.ReadObject(reader), List(Of String))
     For Each lng As String In lngs
      res.Add(New CultureInfo(lng))
     Next
    End Using
   Catch ex As Exception
   End Try
   _SupportedLanguages = res
   Return True
  End Function
#End Region

#Region " Public Methods "
  Public ReadOnly Property SupportedLanguages As List(Of CultureInfo) Implements ITranslationService.SupportedLanguages
   Get
    Return _SupportedLanguages
   End Get
  End Property

  Public Async Function Translate(entriesToTranslate As Dictionary(Of String, String), targetLocale As CultureInfo) As Task(Of Dictionary(Of String, String)) Implements ITranslationService.Translate

   Dim res As New Dictionary(Of String, String)
   Dim _Authentication = New AzureDataMarket(_bingClientId, _bingClientSecret)
   Dim m_Token As AzureDataMarket.Token = Await _Authentication.GetTokenAsync()
   Dim auth As String = m_Token.Header
   For Each ett As KeyValuePair(Of String, String) In entriesToTranslate
    Dim client As New HttpClient()
    Dim strSource As String = ett.Value
    client.DefaultRequestHeaders.Add("Authorization", auth)
    Dim RequestUri As String = "http://api.microsofttranslator.com/v2/Http.svc/Translate?text=" + Net.WebUtility.UrlEncode(strSource) + "&to=" + targetLocale.TwoLetterISOLanguageName
    Try
     Dim strTranslated As String = Await client.GetStringAsync(RequestUri)
     Dim xTranslation As XDocument = XDocument.Parse(strTranslated)
     Dim strTransText As String = xTranslation.Root.FirstNode.ToString()
     res.Add(ett.Key, strTransText)
    Catch ex As Exception
     res.Add(ett.Key, "")
    End Try
   Next
   Return res

  End Function
#End Region

 End Class
End Namespace
