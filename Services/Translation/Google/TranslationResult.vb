Imports Newtonsoft.Json

Namespace Services.Translation.Google
  Public Class TranslationResult

    <JsonProperty("data")>
    Public Property Data As TranslationData

  End Class

  Public Class TranslationData

    <JsonProperty("translations")>
    Public Property Translations As Translation()

  End Class

  Public Class Translation

    <JsonProperty("translatedText")>
    Public Property TranslatedText As String

  End Class
End Namespace
