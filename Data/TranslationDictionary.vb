Imports System.Xml
Imports DotNetNuke.Translator.Common

Namespace Data
 Public Class TranslationDictionary
  Inherits Dictionary(Of String, String)

  Public Property Locale As String = ""
  Public Property FilePath As String = ""

  Public Sub New(locale As String)
   MyBase.New()
   Me.Locale = locale
  End Sub

  Public Sub New(locale As String, filepath As String)
   MyBase.New()
   Me.Locale = locale
   Me.FilePath = filepath
   ReadFile(Me.FilePath)
  End Sub

  Public Sub ReadFile(filepath As String)

   If IO.File.Exists(filepath) Then
    Dim xmlIn As New XmlDocument
    xmlIn.Load(filepath)
    Locale = xmlIn.SelectSingleNode("/dictionary/@locale").InnerText
    For Each entryNode As XmlNode In xmlIn.SelectNodes("/dictionary")
     Me(entryNode.SelectSingleNode("entry").InnerText) = entryNode.SelectSingleNode("value").InnerText
    Next
   End If

  End Sub

  Public Sub Save()
   Save(Me.FilePath)
  End Sub
  Public Sub Save(filepath As String)

   Dim xmlOut As New XmlDocument
   xmlOut.AppendChild(xmlOut.CreateXmlDeclaration("1.0", Nothing, Nothing))
   Dim dicNode As XmlNode = xmlOut.DocumentElement.CreateAndAppendElement("dictionary")
   dicNode.AppendAttribute("locale", Locale)
   For Each kvp As KeyValuePair(Of String, String) In Me
    Dim entryNode As XmlNode = dicNode.CreateAndAppendElement("entry")
    entryNode.CreateAndAppendElement("key", kvp.Key)
    entryNode.CreateAndAppendElement("value", kvp.Value)
   Next
   xmlOut.Save(Me.FilePath)

  End Sub

 End Class
End Namespace
