Imports System.Security
Imports System.Net
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
Imports System.Text

Namespace Common.LEService
 Public Class LEService

  Private _url As String
  Private _accessKey As String
  Public Property IsInError As Boolean = False
  Public Property LastErrorMessage As String = ""
  Public Property LastStatusCode As HttpStatusCode = HttpStatusCode.NoContent
  Public Property Cookies As CookieCollection
  Public Property ModuleId As Integer = -1
  Public Property TabId As Integer = -1

  Public Sub New(url As String, accessKey As SecureString)
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
   _accessKey = Common.Globals.ToInsecureString(accessKey)
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
   objectName = objectName.Replace(" ", "-")
   Return JsonConvert.DeserializeObject(Of List(Of TextInfo))(GetRequest(String.Format("Object/{0}/Version/{1}/Resources?locale={2}", System.Web.HttpUtility.UrlEncode(objectName), objectVersion, targetLocale)))
  End Function

  Public Function GetResourceFile(objectName As String, objectVersion As String, targetLocale As String, fileKey As String) As List(Of TextInfo)
   objectName = objectName.Replace(" ", "-")
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
   If Not String.IsNullOrEmpty(_accessKey) Then
    Dim userId As String = _accessKey.Substring(0, _accessKey.IndexOf("-"))
    wr.Headers.Add("AccessKey", userId)
    Dim salt As String = System.Guid.NewGuid.ToString
    wr.Headers.Add("AccessSalt", salt)
    Dim guid As String = _accessKey.Substring(_accessKey.IndexOf("-") + 1)
    Dim hash As New SHA256Managed
    Dim result As String = Convert.ToBase64String(hash.ComputeHash(ASCIIEncoding.ASCII.GetBytes(guid & salt)))
    wr.Headers.Add("AccessHash", result)
   End If
   Dim responseString As String = DoRequest(wr)
   Return responseString
  End Function

  Private Function PostRequest(relativeUrl As String, body As String) As String
   Dim wr As HttpWebRequest = CType(WebRequest.Create(_url & AddTabAndModule(relativeUrl)), HttpWebRequest)
   If Not String.IsNullOrEmpty(_accessKey) Then
    Dim userId As String = _accessKey.Substring(0, _accessKey.IndexOf("-"))
    wr.Headers.Add("AccessKey", userId)
    Dim salt As String = System.Guid.NewGuid.ToString
    wr.Headers.Add("AccessSalt", salt)
    Dim guid As String = _accessKey.Substring(_accessKey.IndexOf("-") + 1)
    Dim hash As New SHA256Managed
    Dim result As String = Convert.ToBase64String(hash.ComputeHash(ASCIIEncoding.ASCII.GetBytes(guid & salt & body)))
    wr.Headers.Add("AccessHash", result)
   End If
   wr.Method = "POST"
   wr.ContentType = "application/x-www-form-urlencoded"
   body = "body=" & body
   Dim byteArray As Byte() = Text.Encoding.UTF8.GetBytes(body)
   wr.ContentLength = byteArray.Length
   Using dataStream As IO.Stream = wr.GetRequestStream()
    dataStream.Write(byteArray, 0, byteArray.Length)
   End Using
   Dim responseString As String = DoRequest(wr)
   Return responseString
  End Function

  Private Function DoRequest(wr As HttpWebRequest) As String

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
