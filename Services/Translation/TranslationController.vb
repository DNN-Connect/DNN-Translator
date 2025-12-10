Namespace Services.Translation
    Public Class TranslationController

        Public Shared Function TranslationService(appSettings As Common.TranslatorSettings) As ITranslationService
            Select Case appSettings.TranslationProvider.ToLower
                Case "bing"
                    Return New Bing.BingTranslationService(appSettings)
                Case "google"
                    Return New Google.GoogleTranslationService(appSettings)
                Case "azure"
                    Return New Azure.AzureTranslationService(appSettings)
                Case Else
                    Return Nothing
            End Select

        End Function

    End Class
End Namespace