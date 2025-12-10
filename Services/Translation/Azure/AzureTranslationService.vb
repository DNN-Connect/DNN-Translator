Imports System.ServiceModel.Channels
Imports System.ServiceModel
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Xml
Imports Newtonsoft.Json

Namespace Services.Translation.Azure
    Public Class AzureTranslationService
        Implements ITranslationService

        Private _azureApiKey As String = ""
        Private _azureRegion As String = ""
        Private _SupportedLanguages As List(Of CultureInfo)

#Region " Contructors "
        Public Sub New(appSettings As Common.TranslatorSettings)
            _azureApiKey = appSettings.AzureApiKey
            _azureRegion = appSettings.AzureRegion

            If appSettings.AzureLastLanguagesRetrieve < Now.AddMonths(-1) OrElse appSettings.AzureLanguages = "" Then
                Task.Run(Function() Me.RetrieveLanguages).Wait()
                If _SupportedLanguages.Count > 0 Then
                    Dim _cache As String = ""
                    For Each ci As CultureInfo In _SupportedLanguages
                        _cache &= ci.Name & ","
                    Next
                    appSettings.AzureLanguages = _cache.TrimEnd(","c)
                    appSettings.AzureLastLanguagesRetrieve = Now
                    appSettings.Save()
                End If
            Else
                _SupportedLanguages = New List(Of CultureInfo)
                For Each l As String In appSettings.AzureLanguages.Split(","c)
                    Try
                        _SupportedLanguages.Add(New CultureInfo(l))
                    Catch ex As Exception
                    End Try
                Next
            End If

        End Sub
#End Region

#Region " Private Methods "

        Private Async Function RetrieveLanguages() As Task(Of Boolean)
            Dim res As New List(Of CultureInfo)
            Dim resultString As String
            Try
                Using client As New HttpClient()
                    client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", _azureApiKey)
                    client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Region", _azureRegion)
                    Dim result = Await client.GetAsync(New Uri("https://api.cognitive.microsofttranslator.com/languages?api-version=3.0&scope=translation")).ConfigureAwait(False)
                    resultString = Await result.Content.ReadAsStringAsync()
                    Dim translationResult As Azure.LanguageResult = JsonConvert.DeserializeObject(Of Azure.LanguageResult)(resultString)
                    For Each language In translationResult.Languages
                        res.Add(New CultureInfo(language.Key))
                    Next language
                End Using
            Catch ex As Exception
                MsgBox(String.Format("Error connecting to Azure: {0}", ex.Message), MsgBoxStyle.Critical)
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
            For Each ett As KeyValuePair(Of String, String) In entriesToTranslate
                Dim strSource As String = ett.Value
                Dim strTransText As String = String.Empty
                Dim response As HttpResponseMessage
                Try
                    Dim body As Object() = New Object() {New With {Key .Text = strSource}}
                    Dim requestBody = JsonConvert.SerializeObject(body)
                    Using client As New HttpClient()
                        Using request As New HttpRequestMessage()
                            request.Method = HttpMethod.Post
                            request.RequestUri = New Uri("https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=en&to=" + targetLocale.TwoLetterISOLanguageName)
                            request.Headers.Add("Ocp-Apim-Subscription-Key", _azureApiKey)
                            request.Headers.Add("Ocp-Apim-Subscription-Region", _azureRegion)
                            request.Content = New StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                            response = Await client.SendAsync(request).ConfigureAwait(False)
                            strTransText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                            ' Credit for the JSON object decoding :) : https://stackoverflow.com/questions/55752906/json-deserialization-error-with-azure-translation-services
                            Dim translationResult = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, List(Of Dictionary(Of String, String)))))(strTransText)
                            Dim translation = translationResult(0)("translations")(0)("text")
                            res.Add(ett.Key, Translation)
                        End Using
                    End Using

                Catch ex As Exception
                    MsgBox(String.Format("Error connecting to Azure: {0}", ex.Message), MsgBoxStyle.Critical)
                End Try
            Next
            Return res

        End Function
#End Region

    End Class
End Namespace
