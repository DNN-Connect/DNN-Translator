Imports System.Xml
Imports DotNetNuke.Translator.Common

Namespace Data
 Public Class TranslationDictionary
  Inherits Dictionary(Of String, String)

  'Public Property Locale As String = ""
  Public Property FilePath As String = ""

  'Public Sub New(locale As String)
  ' MyBase.New()
  ' 'Me.Locale = locale
  'End Sub

  'Public Sub New(locale As String, filepath As String)
  ' MyBase.New()
  ' 'Me.Locale = locale
  ' Me.FilePath = filepath
  ' ReadFile(Me.FilePath)
  'End Sub

  Public Sub New(filepath As String)
   MyBase.New()
   Me.FilePath = filepath
   ReadFile(Me.FilePath)
  End Sub

  'Public Sub SetTerm(original As String, translation As String)
  ' Me(original) = translation
  'End Sub
  Public Function TranslateTerms(terms As Dictionary(Of String, String)) As Dictionary(Of String, String)
   Dim res As New Dictionary(Of String, String)
   For Each kvp As KeyValuePair(Of String, String) In terms
    If Me.ContainsKey(kvp.Value) Then
     res.Add(kvp.Key, Me(kvp.Value))
    End If
   Next
   Return res
  End Function

  Public Sub ReadFile(filepath As String)

   If IO.File.Exists(filepath) Then
    Dim xmlIn As New XmlDocument
    xmlIn.Load(filepath)
    'Locale = xmlIn.SelectSingleNode("/dictionary/@locale").InnerText
    For Each entryNode As XmlNode In xmlIn.SelectNodes("/dictionary/entry")
     Me(entryNode.SelectSingleNode("key").InnerText) = entryNode.SelectSingleNode("value").InnerText
    Next
   End If

  End Sub

  Public Sub Save()
   Save(Me.FilePath)
  End Sub
  Public Sub Save(filepath As String)

   Dim xmlOut As New XmlDocument
   xmlOut.AppendChild(xmlOut.CreateXmlDeclaration("1.0", Nothing, Nothing))
   Dim dicNode As XmlNode = xmlOut.CreateElement("dictionary")
   xmlOut.AppendChild(dicNode)
   'dicNode.AppendAttribute("locale", Locale)
   For Each kvp As KeyValuePair(Of String, String) In Me
    Dim entryNode As XmlNode = dicNode.CreateAndAppendElement("entry")
    entryNode.CreateAndAppendElement("key", kvp.Key)
    entryNode.CreateAndAppendElement("value", kvp.Value)
   Next
   xmlOut.Save(Me.FilePath)

  End Sub

 End Class
End Namespace
