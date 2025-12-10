Imports Newtonsoft.Json

Namespace Services.Translation.Azure

    <JsonObject()>
    Public Class LanguageResult
        <JsonProperty("translation")>
        Public Property Languages As Dictionary(Of String, Language)
    End Class
    <JsonObject()>
    Public Class Language
        <JsonProperty("name")>
        Public Property Name As String
        <JsonProperty("nativeName")>
        Public Property NativeName As String
        <JsonProperty("dir")>
        Public Property Direction As String
    End Class

End Namespace





