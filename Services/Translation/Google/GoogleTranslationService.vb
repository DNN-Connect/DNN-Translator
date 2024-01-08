Imports System.Net
Imports System.Web
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json

Namespace Services.Translation.Google
  Public Class GoogleTranslationService
    Implements Services.Translation.ITranslationService

    Private _googleApiKey As String = ""
    Private _validGoogleCultureCodes As List(Of String) = New List(Of String)(New String() {"af", "ar", "be", "bg", "ca", "cs", "cy", "da", "de", "el", "en", "es", "et", "fa", "fi", "fr", "ga", "gl", "hi", "hr", "ht", "hu", "id", "is", "it", "iw", "ja", "lt", "lv", "mk", "ms", "mt", "nl", "no", "pl", "pt", "ro", "ru", "sk", "sl", "sq", "sr", "sv", "sw", "th", "tl", "tr", "uk", "vi", "yi", "zh-CN", "zh-TW"})
    Private _SupportedLanguages As New List(Of CultureInfo)

    Public Sub New(appSettings As Common.TranslatorSettings)
      _googleApiKey = appSettings.GoogleApiKey
      For Each l As String In _validGoogleCultureCodes
        Try
          _SupportedLanguages.Add(New System.Globalization.CultureInfo(l))
        Catch ex As Exception
        End Try
      Next
    End Sub

    Public ReadOnly Property SupportedLanguages As System.Collections.Generic.List(Of System.Globalization.CultureInfo) Implements ITranslationService.SupportedLanguages
      Get
        Return _SupportedLanguages
      End Get
    End Property

    Public Async Function Translate(entriesToTranslate As System.Collections.Generic.Dictionary(Of String, String), targetLocale As System.Globalization.CultureInfo) As System.Threading.Tasks.Task(Of System.Collections.Generic.Dictionary(Of String, String)) Implements ITranslationService.Translate

      Dim res As New Dictionary(Of String, String)
      If _googleApiKey = "" Then Return res

      Try
        For Each entry As KeyValuePair(Of String, String) In entriesToTranslate
          Dim client As WebClient = New WebClient()
          client.Encoding = System.Text.Encoding.UTF8
          Dim url As String = "https://www.googleapis.com/language/translate/v2?key=" & _googleApiKey & "&q=" & HttpUtility.UrlEncode(entry.Value) & "&source=en&target=" & targetLocale.Name
          Dim resultString As String = client.DownloadString(url)
          Dim result As TranslationResult = JsonConvert.DeserializeObject(Of TranslationResult)(resultString)
          Dim translatedText As String = HttpUtility.HtmlDecode(result.Data.Translations(0).TranslatedText)
          ' Convert escape unicode characters \uXXXX to normal character
          Dim convertUnicodeRegex As Regex = New Regex("\\[uU]([0-9A-F]{4})", RegexOptions.IgnoreCase)
          translatedText = convertUnicodeRegex.Replace(translatedText, Function(match As Match) ChrW(Integer.Parse(match.Groups(1).Value, NumberStyles.HexNumber)).ToString())
          ' Google doesn't maintain title case very well.
          If Regex.Match(entry.Value, "^[A-Z]").Success And Regex.Match(translatedText, "^[a-z]").Success Then
            If translatedText.Length = 1 Then
              translatedText = translatedText.Substring(0, 1).ToUpper()
            Else
              translatedText = translatedText.Substring(0, 1).ToUpper() & translatedText.Substring(1)
            End If
          End If
          ' Google accidentally converts {0:c} string formats to {0: c}.
          ' We need to remove the extra space
          translatedText = Regex.Replace(translatedText, "{(\d+:) (.+?)}", "{$1$2}")
          res.Add(entry.Key, translatedText)
        Next
      Catch ex As Exception
        MsgBox(String.Format("Error connecting to Google: {0}", ex.Message), MsgBoxStyle.Critical)
      End Try

      Return res

    End Function

  End Class
End Namespace
