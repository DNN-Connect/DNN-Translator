' --------------------------------------------------------------------------------------------------------------------
' <copyright company="Microsoft Corporation" file="AdmAuthentication.cs">
'   Copyright © Microsoft Corporation.  All Rights Reserved.  
' This code released under the terms of the 
' Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)
' </copyright>
' <summary>
'   Authentication helper for the Microsoft Translator on Azure.
' </summary>
' --------------------------------------------------------------------------------------------------------------------

Imports System.Net
Imports System.Runtime.Serialization.Json
Imports System.Text
Imports System.Web

Namespace Services.Translation.Bing

 ''' <summary>
 ''' Authentication helper for the Microsoft Translator on Azure
 ''' </summary>
 Public Class AdmAuthentication
  ''' <summary>
  ''' Url for authentication in Microsoft Translator
  ''' </summary>
  Public Shared ReadOnly DatamarketAccessUri As String = "https://datamarket.accesscontrol.windows.net/v2/OAuth2-13"

  ''' <summary>
  ''' request string needed to request access token
  ''' </summary>
  Private ReadOnly request As String

  ''' <summary>
  ''' Initializes a new instance of the <see cref="AdmAuthentication"/> class.
  ''' </summary>
  ''' <param name="clientId">The client id.</param>
  ''' <param name="clientSecret">The client secret.</param>
  Public Sub New(clientId As String, clientSecret As String)
   ' If clientid or client secret has special characters, encode before sending request
   Me.request = String.Format("grant_type=client_credentials&client_id={0}&client_secret={1}&scope=http://api.microsofttranslator.com", HttpUtility.UrlEncode(clientId), HttpUtility.UrlEncode(clientSecret))
  End Sub

  ''' <summary>
  ''' Requests the access token
  ''' </summary>
  ''' <returns>The access token</returns>
  Public Function GetAccessToken() As AdmAccessToken
   Return HttpPost(DatamarketAccessUri, Me.request)
  End Function

  ''' <summary>
  ''' Helper method to formulate the proper request
  ''' </summary>
  ''' <param name="datamarketAccessUri">Web Address for the request</param>
  ''' <param name="requestDetails">request details</param>
  ''' <returns>The access toen</returns>
  Private Shared Function HttpPost(datamarketAccessUri As String, requestDetails As String) As AdmAccessToken
   ' Prepare OAuth request 
   Dim webRequest__1 = WebRequest.Create(datamarketAccessUri)
   webRequest__1.ContentType = "application/x-www-form-urlencoded"
   webRequest__1.Method = "POST"
   Dim bytes = Encoding.ASCII.GetBytes(requestDetails)
   webRequest__1.ContentLength = bytes.Length
   Using outputStream = webRequest__1.GetRequestStream()
    outputStream.Write(bytes, 0, bytes.Length)
   End Using

   Using webResponse = webRequest__1.GetResponse()
    Dim serializer = New DataContractJsonSerializer(GetType(AdmAccessToken))

    ' Get deserialized object from JSON stream
    Dim token = DirectCast(serializer.ReadObject(webResponse.GetResponseStream()), AdmAccessToken)
    Return token
   End Using
  End Function
 End Class
End Namespace
