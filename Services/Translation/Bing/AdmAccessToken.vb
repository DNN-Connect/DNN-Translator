' --------------------------------------------------------------------------------------------------------------------
' <copyright company="Microsoft Corporation" file="AdmAccessToken.cs">
'   Copyright © Microsoft Corporation.  All Rights Reserved.  
' This code released under the terms of the 
' Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)
' </copyright>
' <summary>
'   Authentication token data class for the Microsoft Translator on Azure.
' </summary>
' --------------------------------------------------------------------------------------------------------------------

Imports System.Runtime.Serialization

Namespace Services.Translation.Bing

 ''' <summary>
 ''' Data cclass capturing an access token
 ''' </summary>
 <DataContract()> _
 Public Class AdmAccessToken
  ''' <summary>
  ''' Gets or sets the access token
  ''' </summary>
  <DataMember(Name:="access_token")> _
  Public Property AccessToken() As String
   Get
    Return m_AccessToken
   End Get
   Set(value As String)
    m_AccessToken = value
   End Set
  End Property
  Private m_AccessToken As String

  ''' <summary>
  ''' Gets or sets the token type
  ''' </summary>
  <DataMember(Name:="token_type")> _
  Public Property TokenType() As String
   Get
    Return m_TokenType
   End Get
   Set(value As String)
    m_TokenType = value
   End Set
  End Property
  Private m_TokenType As String

  ''' <summary>
  ''' Gets or sets the expiration of the access token
  ''' </summary>
  <DataMember(Name:="expires_in")> _
  Public Property ExpiresIn() As String
   Get
    Return m_ExpiresIn
   End Get
   Set(value As String)
    m_ExpiresIn = value
   End Set
  End Property
  Private m_ExpiresIn As String

  ''' <summary>
  ''' Gets or sets the scope of access token
  ''' </summary>
  <DataMember(Name:="scope")> _
  Public Property Scope() As String
   Get
    Return m_Scope
   End Get
   Set(value As String)
    m_Scope = value
   End Set
  End Property
  Private m_Scope As String

  Public Property Expires As DateTime = DateTime.MinValue

 End Class
End Namespace
