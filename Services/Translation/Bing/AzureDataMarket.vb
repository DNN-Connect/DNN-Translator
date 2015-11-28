Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace Services.Translation.Bing
 Public Class AzureDataMarket
  ReadOnly m_ClientId As String
  ReadOnly m_ClientSecret As String
  ReadOnly m_Request As String
  ReadOnly DATAMARKET_ACCESS_URI As String = "https://datamarket.accesscontrol.windows.net/v2/OAuth2-13"

  Public Sub New(clientId As String, clientSecret As String)
   Me.m_ClientId = Uri.EscapeDataString(clientId.Trim)
   Me.m_ClientSecret = Uri.EscapeDataString(clientSecret.Trim)
   Me.m_Request = String.Format("grant_type=client_credentials&client_id={0}&client_secret={1}&scope=http://api.microsofttranslator.com", m_ClientId, m_ClientSecret)
  End Sub

  Public Async Function GetTokenAsync() As Task(Of Token)
   ' build request
   Dim _Request = System.Net.WebRequest.Create(DATAMARKET_ACCESS_URI)
   _Request.ContentType = "application/x-www-form-urlencoded"
   _Request.Method = "POST"

   ' make request
   Dim _Bytes = Encoding.UTF8.GetBytes(Me.m_Request)
   Using _Stream = Await _Request.GetRequestStreamAsync()
    _Stream.Write(_Bytes, 0, _Bytes.Length)
   End Using

   ' deserialize response
   Try
    Using _Response = Await _Request.GetResponseAsync()
     Dim _Serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(Token))
     Dim _Token = Await Task.Run(Function() DirectCast(_Serializer.ReadObject(_Response.GetResponseStream()), Token))
     Return _Token
    End Using
   Catch generatedExceptionName As Exception
    System.Diagnostics.Debugger.Break()
    Throw
   End Try
  End Function

  <System.Runtime.Serialization.DataContract>
  Public Class Token
   <System.Runtime.Serialization.DataMember>
   Public Property access_token() As String
    Get
     Return m_access_token
    End Get
    Set
     m_access_token = Value
    End Set
   End Property
   Private m_access_token As String

   <System.Runtime.Serialization.DataMember>
   Public Property token_type() As String
    Get
     Return m_token_type
    End Get
    Set
     m_token_type = Value
    End Set
   End Property
   Private m_token_type As String

   Private m_expires_in As Integer
   <System.Runtime.Serialization.DataMember>
   Public Property expires_in() As Integer
    Get
     Return m_expires_in
    End Get
    Set
     m_expires_in = Value
     ExpirationDate = DateTime.Now.AddSeconds(Value)
    End Set
   End Property

   <System.Runtime.Serialization.DataMember>
   Public Property scope() As String
    Get
     Return m_scope
    End Get
    Set
     m_scope = Value
    End Set
   End Property
   Private m_scope As String

   Private ExpirationDate As DateTime = DateTime.MinValue
   Public ReadOnly Property IsExpired() As Boolean
    Get
     Return ExpirationDate < DateTime.Now
    End Get
   End Property

   Public ReadOnly Property Header() As String
    Get
     Return String.Format("Bearer {0}", access_token)
    End Get
   End Property
  End Class
 End Class

End Namespace

