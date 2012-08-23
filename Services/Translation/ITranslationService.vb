Namespace Services.Translation
 Public Interface ITranslationService

  Function Translate(entriesToTranslate As Dictionary(Of String, String), targetLocale As CultureInfo) As Dictionary(Of String, String)
  ReadOnly Property SupportedLanguages() As List(Of CultureInfo)

 End Interface
End Namespace
