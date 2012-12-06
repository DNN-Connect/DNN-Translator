Imports System.Security
Imports System.Net
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions

Namespace Common.LEService
 Public Class LEService

  Private _url As String
  Private _username As String
  Private _password As SecureString
  Public Property IsInError As Boolean = False
  Public Property LastErrorMessage As String = ""
  Public Property LastStatusCode As HttpStatusCode = HttpStatusCode.NoContent
  Public Property Cookies As CookieCollection
  Public Property ModuleId As Integer = -1
  Public Property TabId As Integer = -1
  Private _antiForgeryTokenName As String = ""
  Public Property AntiForgeryTokenName() As String
   Get
    If String.IsNullOrEmpty(_antiForgeryTokenName) Then
     Try
      _antiForgeryTokenName = ApplicationState.GetValue(Of String)("AntiForgeryTokenName")
     Catch ex As Exception
     End Try
    End If
    Return _antiForgeryTokenName
   End Get
   Set(ByVal value As String)
    If value <> "" Then
     _antiForgeryTokenName = value
     ApplicationState.SetValue("AntiForgeryTokenName", value)
    End If
   End Set
  End Property
  Private _antiForgeryToken As String = ""
  Public Property AntiForgeryToken() As String
   Get
    If String.IsNullOrEmpty(_antiForgeryToken) Then
     Try
      _antiForgeryToken = ApplicationState.GetValue(Of String)("AntiForgeryToken")
     Catch ex As Exception
     End Try
    End If
    Return _antiForgeryToken
   End Get
   Set(ByVal value As String)
    If value <> "" Then
     _antiForgeryToken = value
     ApplicationState.SetValue("AntiForgeryToken", value)
    End If
   End Set
  End Property

  Public Sub New(url As String, username As String, password As SecureString)
   _url = url.ToLower
   If Not _url.StartsWith("http://") And Not _url.StartsWith("https://") Then
    _url = "http://" & _url
   End If
   Dim m As Match = Regex.Match(_url, "(?i)(.*)\?tabid=(\d+)&moduleid=(\d+)(?-i)")
   If m.Success Then
    _url = m.Groups(1).Value & "/"
    TabId = Integer.Parse(m.Groups(2).Value)
    ModuleId = Integer.Parse(m.Groups(3).Value)
   End If
   _username = username
   _password = password
   Try
    Cookies = Common.ApplicationState.GetValue(Of CookieCollection)("Cookies")
   Catch ex As Exception
   End Try
   If Cookies Is Nothing Then Cookies = New CookieCollection
  End Sub

  Public Function GetEditLanguages() As List(Of CultureInfo)
   Dim locales As List(Of String) = JsonConvert.DeserializeObject(Of List(Of String))(GetRequest("EditLocales"))
   Dim res As New List(Of CultureInfo)
   For Each locale As String In locales
    res.Add(New CultureInfo(locale))
   Next
   Return res
  End Function

  Public Function GetResources(objectName As String, objectVersion As String, targetLocale As String, start As String) As List(Of TextInfo)
   Return JsonConvert.DeserializeObject(Of List(Of TextInfo))(GetRequest(String.Format("Object/{0}/Version/{1}/Resources?locale={2}", System.Web.HttpUtility.UrlEncode(objectName), objectVersion, targetLocale)))
  End Function

  Public Function GetResourceFile(objectName As String, objectVersion As String, targetLocale As String, fileKey As String) As List(Of TextInfo)
   Return JsonConvert.DeserializeObject(Of List(Of TextInfo))(GetRequest(String.Format("Object/{0}/Version/{1}/File/{2}?locale={3}", System.Web.HttpUtility.UrlEncode(objectName), objectVersion, fileKey.Replace("_", "=").Replace(".", "-"), targetLocale)))
  End Function

  Public Function UploadResources(texts As List(Of TextInfo)) As String
   Return PostRequest("/UpdateResources", JsonConvert.SerializeObject(texts, New DTConverter))
  End Function

#Region " Web Request "
  Private Function AddTabAndModule(relativeUrl As String) As String
   If TabId <> -1 Then
    If relativeUrl.Contains("?") Then
     relativeUrl &= "&"
    Else
     relativeUrl &= "?"
    End If
    relativeUrl &= "TabId=" & TabId.ToString
   End If
   If ModuleId <> -1 Then
    relativeUrl &= "&ModuleId=" & ModuleId.ToString
   End If
   Return relativeUrl
  End Function

  Private Function GetRequest(relativeUrl As String) As String
   Dim wr As HttpWebRequest = CType(WebRequest.Create(_url & AddTabAndModule(relativeUrl)), HttpWebRequest)
   Dim responseString As String = DoRequest(wr)
   Return responseString
  End Function

  Private Function PostRequest(relativeUrl As String, body As String) As String
   Dim aft As String = GetRequest("aft").Trim(""""c)
   If AntiForgeryTokenName <> "" Then
    If Cookies(AntiForgeryTokenName) IsNot Nothing Then
     If Cookies(AntiForgeryTokenName).Value = "" Then
      Cookies(AntiForgeryTokenName).Value = AntiForgeryToken
     End If
    End If
   End If
   body = "__RequestVerificationToken=" & aft & "&body=" & body
   Dim wr As HttpWebRequest = CType(WebRequest.Create(_url & AddTabAndModule(relativeUrl)), HttpWebRequest)
   wr.Headers.Add("__RequestVerificationToken", aft)
   wr.Method = "POST"
   wr.ContentType = "application/x-www-form-urlencoded"
   Dim byteArray As Byte() = Text.Encoding.UTF8.GetBytes(body)
   wr.ContentLength = byteArray.Length
   Using dataStream As IO.Stream = wr.GetRequestStream()
    dataStream.Write(byteArray, 0, byteArray.Length)
   End Using
   Dim responseString As String = DoRequest(wr)
   Return responseString
  End Function

  Private Function DoRequest(wr As HttpWebRequest) As String

   If Not String.IsNullOrEmpty(_username) Then
    wr.Credentials = New NetworkCredential(_username, _password)
   End If
   wr.CookieContainer = New CookieContainer
   wr.CookieContainer.Add(Cookies)

   Dim responseString As String = ""
   LastErrorMessage = ""
   LastStatusCode = HttpStatusCode.NoContent

   Try
    Using response As HttpWebResponse = CType(wr.GetResponse, HttpWebResponse)
     Using sr As New IO.StreamReader(response.GetResponseStream)
      responseString = sr.ReadToEnd
     End Using
     LastStatusCode = response.StatusCode
     Cookies = response.Cookies
    End Using
   Catch ex As System.Net.WebException
    If ex.Message.Contains("404") Then
     Try
      ' attempt to wake up dnn
      Dim mockWr As WebRequest = WebRequest.Create(Left(_url, _url.ToLower.IndexOf("/desktopmodules")))
      Using mockResponse As WebResponse = mockWr.GetResponse
       ' do nothing with it
      End Using
      ' try again - we have to clone the request
      Dim wrSecondTry As WebRequest = WebRequest.Create(wr.RequestUri)
      wrSecondTry.Credentials = wr.Credentials
      wrSecondTry.Method = wr.Method
      wrSecondTry.ContentType = wr.ContentType
      If wr.Method = "POST" Then
       Using dataStream As IO.Stream = wr.GetRequestStream()
        Dim byteArray(CInt(dataStream.Length)) As Byte
        wr.ContentLength = byteArray.Length
        dataStream.Read(byteArray, 0, CInt(dataStream.Length))
        Using cloneStream As IO.Stream = wrSecondTry.GetRequestStream
         cloneStream.Write(byteArray, 0, byteArray.Length)
        End Using
       End Using
      End If
      Using response As HttpWebResponse = CType(wrSecondTry.GetResponse, HttpWebResponse)
       Using sr As New IO.StreamReader(response.GetResponseStream)
        responseString = sr.ReadToEnd
       End Using
       LastStatusCode = response.StatusCode
       Cookies = response.Cookies
      End Using
     Catch ex2 As Exception
      IsInError = True
      LastErrorMessage = ex2.Message
     End Try
    Else
     IsInError = True
     LastErrorMessage = ex.Message
    End If
   Catch ex1 As Exception
    IsInError = True
    LastErrorMessage = ex1.Message
   End Try

   Common.ApplicationState.SetValue("Cookies", Cookies)

   Return responseString
  End Function
#End Region

 End Class
End Namespace
