Imports System.ServiceModel.Channels
Imports System.ServiceModel

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
    _SupportedLanguages = RetrieveLanguages()
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

  Private Function RetrieveLanguages() As List(Of CultureInfo)

   Dim res As New List(Of CultureInfo)
   If _bingClientId = "" Then Return res

   Try
    ' Add TranslatorService as a service reference, Address:http://api.microsofttranslator.com/V2/Soap.svc
    Dim translator = New MicrosoftTranslatorService.LanguageServiceClient()
    Dim binding = DirectCast(translator.Endpoint.Binding, BasicHttpBinding)
    binding.MaxReceivedMessageSize = Integer.MaxValue
    binding.MaxBufferSize = Integer.MaxValue

    ' Set Authorization header before sending the request
    Dim httpRequestProperty = New HttpRequestMessageProperty() With {.Method = "POST"}
    httpRequestProperty.Headers.Add("Authorization", AccessToken)

    Dim languagesForTranslate As String()

    ' Creates a block within which an OperationContext object is in scope.
    Using New OperationContextScope(translator.InnerChannel)
     OperationContext.Current.OutgoingMessageProperties(HttpRequestMessageProperty.Name) = httpRequestProperty
     ' Keep appId parameter blank as we are sending access token in authorization header.
     languagesForTranslate = translator.GetLanguagesForTranslate(String.Empty)
    End Using

    For Each lang As String In languagesForTranslate
     Try
      res.Add(New CultureInfo(lang))
     Catch ex As Exception
     End Try
    Next
   Catch ex As Exception

   End Try

   Return res

  End Function
#End Region

#Region " Public Methods "
  Public ReadOnly Property SupportedLanguages As System.Collections.Generic.List(Of System.Globalization.CultureInfo) Implements ITranslationService.SupportedLanguages
   Get
    Return _SupportedLanguages
   End Get
  End Property

  Public Function Translate(entriesToTranslate As Dictionary(Of String, String), targetLocale As CultureInfo) As Dictionary(Of String, String) Implements ITranslationService.Translate

   Dim res As New Dictionary(Of String, String)
   If _bingClientId = "" Then Return res

   Try
    ' get bing reference and response
    Dim translator = New MicrosoftTranslatorService.LanguageServiceClient()
    Dim binding = DirectCast(translator.Endpoint.Binding, BasicHttpBinding)
    binding.MaxReceivedMessageSize = Integer.MaxValue
    binding.MaxBufferSize = Integer.MaxValue

    ' Set Authorization header before sending the request
    Dim httpRequestProperty = New HttpRequestMessageProperty() With { _
      .Method = "POST" _
    }
    httpRequestProperty.Headers.Add("Authorization", AccessToken)
    Dim options = New MicrosoftTranslatorService.TranslateOptions()

    Dim requestSize = 0
    Dim lastEntryHash = entriesToTranslate.Last().GetHashCode()
    Dim translationBatch As New Dictionary(Of String, String)

    For Each entry As KeyValuePair(Of String, String) In entriesToTranslate

     translationBatch.Add(entry.Key, entry.Value)
     requestSize += entry.Value.ToString().Length

     ' The total of all texts to be translated must not exceed 10000 characters. The maximum number of array elements is 2000. 
     If (requestSize <= 8000) AndAlso (entriesToTranslate.Count < 2000) AndAlso (entry.GetHashCode() <> lastEntryHash) Then
      Continue For
     End If

     Dim stringsToTranslate = translationBatch.[Select](Function(t) t.Value.ToString()).ToArray()
     Dim translatedTexts As MicrosoftTranslatorService.TranslateArrayResponse() = Nothing
     Try
      ' Creates a block within which an OperationContext object is in scope.
      Using New OperationContextScope(translator.InnerChannel)
       OperationContext.Current.OutgoingMessageProperties(HttpRequestMessageProperty.Name) = httpRequestProperty
       translatedTexts = translator.TranslateArray(String.Empty, stringsToTranslate, "en", targetLocale.Name, options)
      End Using
     Catch e As Exception

     End Try

     ' loop orginal resx and create new one from translated text
     Dim l = 0
     For Each translateEntry As KeyValuePair(Of String, String) In translationBatch
      If (translatedTexts Is Nothing) OrElse (l >= translatedTexts.Length) Then
       ' incomplete
      Else
       res.Add(translateEntry.Key, translatedTexts(l).TranslatedText)
      End If
      l += 1
     Next

     'If (abortRequested IsNot Nothing) AndAlso abortRequested() Then
     ' Exit For
     'End If

     translationBatch = New Dictionary(Of String, String)
     requestSize = 0
    Next
   Catch ex As Exception

   End Try

   Return res

  End Function
#End Region

 End Class
End Namespace
