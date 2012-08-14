Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Net
Imports System.IO
Imports System.Runtime.Serialization.Json
Imports System.Runtime.Serialization
Imports System.Web
Imports System.ServiceModel.Channels
Imports System.ServiceModel
Imports System.Xml
Imports System.Globalization

Namespace Common.Bing
 Public Class Translate

  Private Shared Sub Main(args As String())
   Dim admToken As AdmAccessToken
   Dim headerValue As String
   'Get Client Id and Client Secret from https://datamarket.azure.com/developer/applications/
   'Refer obtaining AccessToken (http://msdn.microsoft.com/en-us/library/hh454950.aspx) 
   Dim admAuth As New AdmAuthentication("clientID", "client secret")
   Try
    admToken = admAuth.GetAccessToken()
    ' Create a header with the access_token property of the returned token
    headerValue = "Bearer " + admToken.AccessToken
    Dim languageCodes As String() = {"en", "fr", "uk"}

    GetLanguageNamesMethod(headerValue, languageCodes)
   Catch e As WebException
    ProcessWebException(e)
    Console.WriteLine("Press any key to continue...")
    Console.ReadKey(True)
   Catch ex As Exception

    Console.WriteLine(ex.Message)
    Console.WriteLine("Press any key to continue...")
    Console.ReadKey(True)
   End Try
  End Sub

  Private Shared Sub GetLanguageNamesMethod(authToken As String, languageCodes As String())
   Dim uri As String = "http://api.microsofttranslator.com/v2/Http.svc/GetLanguageNames?locale=en"
   ' create the request
   Dim request As HttpWebRequest = DirectCast(WebRequest.Create(uri), HttpWebRequest)
   request.Headers.Add("Authorization", authToken)
   request.ContentType = "text/xml"
   request.Method = "POST"
   Dim dcs As New System.Runtime.Serialization.DataContractSerializer(Type.[GetType]("System.String[]"))
   Using stream As System.IO.Stream = request.GetRequestStream()
    dcs.WriteObject(stream, languageCodes)
   End Using
   Dim response As WebResponse = Nothing
   Try
    response = request.GetResponse()

    Using stream As Stream = response.GetResponseStream()
     Dim languageNames As String() = DirectCast(dcs.ReadObject(stream), String())

     Console.WriteLine("Language codes = Language name")
     For i As Integer = 0 To languageNames.Length - 1
      Console.WriteLine("      {0}       = {1}", languageCodes(i), languageNames(i))
     Next
     Console.WriteLine("Press any key to continue...")
     Console.ReadKey(True)
    End Using
   Catch
    Throw
   Finally
    If response IsNot Nothing Then
     response.Close()
     response = Nothing
    End If
   End Try
  End Sub

  Private Shared Sub ProcessWebException(e As WebException)
   Console.WriteLine("{0}", e.ToString())
   ' Obtain detailed error information
   Dim strResponse As String = String.Empty
   Using response As HttpWebResponse = DirectCast(e.Response, HttpWebResponse)
    Using responseStream As Stream = response.GetResponseStream()
     Using sr As New StreamReader(responseStream, System.Text.Encoding.ASCII)
      strResponse = sr.ReadToEnd()
     End Using
    End Using
   End Using
   Console.WriteLine("Http status code={0}, error message={1}", e.Status, strResponse)
  End Sub

  ''' <summary>
  ''' Web address of landing page for Azure market for Microsoft Translator
  ''' </summary>
  Public Const AzureMarketTranslatorLandingPage As String = "https://datamarket.azure.com/dataset/1899a118-d202-492c-aa16-ba21c33c06cb"

  ''' <summary>
  ''' resx file extension
  ''' </summary>
  Private Const ResxExtension As String = ".resx"

  ''' <summary>
  ''' Delegate that is used to receive a log event from the issuer of a translation request
  ''' </summary>
  ''' <param name="message">The message.</param>
  Public Delegate Sub LogEvent(message As String)

  ''' <summary>
  ''' Delegate that is used to issue a abort request in a translation process by the issuer of the request
  ''' </summary>
  ''' <returns>Boolean value indiating whether translation shall be aborted</returns>
  Public Delegate Function ShallAbort() As Boolean

  ''' <summary>
  ''' Find the matching culture from a Resx file name from a number of given cultures
  ''' </summary>
  ''' <param name="candidates">the cultures to choose from</param>
  ''' <param name="fileName">the resx file name</param>
  ''' <returns>The best matching culture</returns>
  Public Shared Function FindMatchingCultureFromResxFileName(candidates As IEnumerable(Of CultureInfo), fileName As String) As CultureInfo
   If ResxExtension.Equals(Path.GetExtension(fileName), StringComparison.CurrentCultureIgnoreCase) Then
    Dim temp = Path.GetFileNameWithoutExtension(fileName)
    Dim cultureExtension = Path.GetExtension(temp)
    If Not String.IsNullOrEmpty(cultureExtension) Then
     Try
      Dim resxCi = CultureInfo.GetCultureInfo(cultureExtension.Trim("."c))
      For Each cInfo As CultureInfo In candidates.Where(Function(ci) (ci.LCID = resxCi.LCID) OrElse (ci.LCID = resxCi.Parent.LCID))
       Return cInfo
      Next
     Catch generatedExceptionName As ArgumentException
     End Try
    End If
   End If

   Return Nothing
  End Function

  ''' <summary>
  ''' Returns the credentials string for an azure account key
  ''' </summary>
  ''' <param name="clientID">The service account client ID</param>
  ''' <param name="clientSecret">The service account client Secret</param>
  ''' <param name="logEvent">Event logging delegate</param>
  ''' <returns>The servcie credientals string</returns>
  Public Shared Function RequestAuthenticationToken(clientID As String, clientSecret As String, logEvent As LogEvent) As String
   Dim admAuth = New AdmAuthentication(clientID, clientSecret)
   Try
    Dim admToken = admAuth.GetAccessToken()
    Dim tokenReceived = DateTime.Now

    ' Create a header with the access_token property of the returned token
    Return "Bearer " & admToken.AccessToken
   Catch e As Exception
    LogExceptionError(e, logEvent)
    Return Nothing
   End Try
  End Function

  ''' <summary>
  ''' Requests th eids for all language supported by the service
  ''' </summary>
  ''' <param name="authenticationToken">The service credentials, e.g. bing app id or Azure Account key.</param>
  ''' <param name="logEvent">Event logging delegate</param>
  ''' <returns>Array of identifiers of supported languages</returns>
  Public Shared Function GetLanguagesForTranslate(authenticationToken As String, logEvent As LogEvent) As String()
   ' Add TranslatorService as a service reference, Address:http://api.microsofttranslator.com/V2/Soap.svc
   Dim translator = New MicrosoftTranslatorService.LanguageServiceClient()
   Dim binding = DirectCast(translator.Endpoint.Binding, BasicHttpBinding)
   binding.MaxReceivedMessageSize = Integer.MaxValue
   binding.MaxBufferSize = Integer.MaxValue

   ' Set Authorization header before sending the request
   Dim httpRequestProperty = New HttpRequestMessageProperty() With { _
     .Method = "POST" _
   }
   httpRequestProperty.Headers.Add("Authorization", authenticationToken)

   Dim languagesForTranslate As String()

   ' Creates a block within which an OperationContext object is in scope.
   Using New OperationContextScope(translator.InnerChannel)
    OperationContext.Current.OutgoingMessageProperties(HttpRequestMessageProperty.Name) = httpRequestProperty

    ' Keep appId parameter blank as we are sending access token in authorization header.
    languagesForTranslate = translator.GetLanguagesForTranslate(String.Empty)
   End Using

   Return languagesForTranslate
  End Function

  ''' <summary>Translates resx files in directory recursive</summary>
  ''' <param name="searchRoot">The search root path.</param>
  ''' <param name="searchFilter">The search filter.</param>        
  ''' <param name="fromCultureBing">The from culture.</param>
  ''' <param name="toCultureBing">The to culture.</param>
  ''' <param name="toCultureFile">The Target culture for the resx file.</param>
  ''' <param name="authenticationToken">The service credentials, e.g. bing app id or Azure Account key.</param>
  ''' <param name="logEvent">Event logging delegate</param>
  ''' <param name="abortRequested">Event for checking whether abort has been requested</param>
  ''' <returns>Number of files translated</returns>
  Public Shared Function TranslateResXDirectory(searchRoot As String, searchFilter As String, fromCultureBing As CultureInfo, toCultureBing As CultureInfo, toCultureFile As CultureInfo, authenticationToken As String, _
   logEvent As LogEvent, abortRequested As ShallAbort) As Integer
   If (abortRequested IsNot Nothing) AndAlso abortRequested() Then
    Return 0
   End If

   If Not Directory.Exists(searchRoot) Then
    Return 0
   End If

   ' Recurse directories
   Dim subDirs = Directory.GetDirectories(searchRoot)
   Dim filesTranslated = 0
   For Each dir As String In subDirs
    filesTranslated += TranslateResXDirectory(dir, searchFilter, fromCultureBing, toCultureBing, toCultureFile, authenticationToken, _
     logEvent, abortRequested)
    If (abortRequested IsNot Nothing) AndAlso abortRequested() Then
     Exit For
    End If
   Next

   ' Translate resx files
   If (abortRequested Is Nothing) OrElse Not abortRequested() Then
    Dim files = Directory.GetFiles(searchRoot, searchFilter)
    filesTranslated += files.Count(Function(file) TranslateResXFile(file, fromCultureBing, toCultureBing, toCultureFile, authenticationToken, logEvent, _
     abortRequested))
   End If

   Return filesTranslated
  End Function

  ''' <summary>Translates resx file</summary>
  ''' <param name="fromResxFilePath">The from resx path.</param>
  ''' <param name="fromCultureBing">The Source culture in Bing.</param>
  ''' <param name="toCultureBing">The Target culture in Bing.</param>
  ''' <param name="toCultureFile">The Target culture for the resx file.</param>
  ''' <param name="authenticationToken">The service credentials, e.g. bing app id or Azure Account Key.</param>
  ''' <param name="logEvent">Event logging delegate</param>
  ''' <param name="abortRequested">Event for checking whether abort has been requested</param>
  ''' <returns>True if a file has ben translated</returns>
  Public Shared Function TranslateResXFile(fromResxFilePath As String, fromCultureBing As CultureInfo, toCultureBing As CultureInfo, toCultureFile As CultureInfo, authenticationToken As String, logEvent As LogEvent, _
   abortRequested As ShallAbort) As Boolean
   If (abortRequested IsNot Nothing) AndAlso abortRequested() Then
    Return False
   End If

   If Not File.Exists(fromResxFilePath) Then
    Return False
   End If

   Dim resXResourceReader = New System.Resources.ResXResourceReader(fromResxFilePath)

   ' construct To RESX path
   Dim directory = Path.GetDirectoryName(fromResxFilePath)
   Dim resXFromFileExtension = Path.GetExtension(fromResxFilePath)
   Dim resXRootFileName = Path.GetFileNameWithoutExtension(fromResxFilePath)
   Dim origLangTag = Path.GetExtension(resXRootFileName)
   Dim origCulture As CultureInfo = Nothing
   Try
    If Not String.IsNullOrEmpty(origLangTag) Then
     origCulture = CultureInfo.GetCultureInfo(origLangTag.Trim("."c))
    End If
   Catch generatedExceptionName As ArgumentException
   End Try

   Dim toResxPath As String

   If origCulture IsNot Nothing Then
    toResxPath = Path.ChangeExtension(resXRootFileName, toCultureFile.Name & resXFromFileExtension)
   Else
    toResxPath = resXRootFileName & "." & toCultureFile.Name & resXFromFileExtension
   End If

   'RaiseEvent LogEvent(toResxPath & " ")

   If Not String.IsNullOrEmpty(directory) Then
    toResxPath = Path.Combine(directory, toResxPath)
   End If

   ' get To Resource Writer
   Dim toResxStream As Stream = New FileStream(toResxPath, FileMode.OpenOrCreate)
   Dim toResxWriter = New System.Resources.ResXResourceWriter(toResxStream)
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
    httpRequestProperty.Headers.Add("Authorization", authenticationToken)

    Dim options = New MicrosoftTranslatorService.TranslateOptions()
    Dim resXEntryList = resXResourceReader.Cast(Of DictionaryEntry)().ToList()

    Dim entriesToTranslate = New List(Of DictionaryEntry)()
    Dim requestSize = 0
    Dim lastEntryHash = resXEntryList.Last().GetHashCode()
    For Each entry As DictionaryEntry In resXEntryList
     entriesToTranslate.Add(entry)
     requestSize += entry.Value.ToString().Length

     ' The total of all texts to be translated must not exceed 10000 characters. The maximum number of array elements is 2000. 
     If (requestSize <= 8000) AndAlso (entriesToTranslate.Count < 2000) AndAlso (entry.GetHashCode() <> lastEntryHash) Then
      Continue For
     End If

     Dim stringsToTranslate = entriesToTranslate.[Select](Function(t) t.Value.ToString()).ToArray()
     Dim translatedTexts As MicrosoftTranslatorService.TranslateArrayResponse()
     Try
      ' Creates a block within which an OperationContext object is in scope.
      Using New OperationContextScope(translator.InnerChannel)
       OperationContext.Current.OutgoingMessageProperties(HttpRequestMessageProperty.Name) = httpRequestProperty
       translatedTexts = translator.TranslateArray(String.Empty, stringsToTranslate, fromCultureBing.Name, toCultureBing.Name, options)
      End Using
     Catch e As Exception
      LogExceptionError(e, logEvent)
      Return False
     End Try

     'RaiseEvent LogEvent(".")

     ' loop orginal resx and create new one from translated text
     Dim l = 0
     For Each translateEntry As DictionaryEntry In entriesToTranslate
      If (translatedTexts Is Nothing) OrElse (l >= translatedTexts.Length) Then
       'RaiseEvent LogEvent("inclomplete Translation")

       Exit For
      End If

      Dim translatedText = translatedTexts(l).TranslatedText
      toResxWriter.AddResource(translateEntry.Key.ToString(), translatedText)
      l += 1
     Next

     'Application.DoEvents()

     If (abortRequested IsNot Nothing) AndAlso abortRequested() Then
      Exit For
     End If

     entriesToTranslate.Clear()
     requestSize = 0
    Next
   Finally
    ' close RESX
    toResxWriter.Close()
    toResxStream.Close()
    'RaiseEvent LogEvent(vbCr & vbLf)
   End Try

   Return True
  End Function

  ''' <summary>
  ''' General exception handling
  ''' </summary>
  ''' <param name="ex">The exception</param>
  ''' <param name="logEvent">Loging delegate</param>
  Private Shared Sub LogExceptionError(ex As Exception, logEvent As LogEvent)
   If TypeOf ex Is ArgumentOutOfRangeException Then
    If logEvent IsNot Nothing Then
     logEvent(vbCr & vbLf)
     logEvent("Error: " & ex.Message)
    End If
   ElseIf TypeOf ex Is Services.Protocols.SoapHeaderException Then
    If logEvent IsNot Nothing Then
     logEvent(vbCr & vbLf)
     logEvent("Microsoft Translator API Error: " & ex.Message)
    End If
   Else
    If logEvent IsNot Nothing Then
     logEvent(vbCr & vbLf)
     logEvent("General Error: " & ex.Message)
    End If
   End If
  End Sub

 End Class
End Namespace

