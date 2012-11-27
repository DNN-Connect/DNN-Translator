Imports System.Linq
Imports System.Xml

Namespace Common
 Public Class ResourceFile
  Inherits XmlDocument

  Public Property Resources As New Dictionary(Of String, Resource)
  Public Property FilePath As String = ""
  Public Property fileKey As String = ""
  Public Property Copyright As String = ""

  Public Sub GetResource(key As String, ByRef value As String)
   If Resources.ContainsKey(key) Then
    value = Resources(key).Value
   End If
  End Sub

  Public Sub SetResourceValue(key As String, value As String)
   SetResourceValue(key, value, Now)
  End Sub

  Public Sub SetResourceValue(key As String, value As String, lastModified As DateTime)
   If value = "" AndAlso Resources.ContainsKey(key) Then
    Resources.Remove(key)
    Exit Sub
   End If
   If Resources.ContainsKey(key) Then
    If Resources(key).Value <> value Then
     Resources(key).Value = value
     Resources(key).LastModified = lastModified
    End If
   Else
    Dim r As New Resource With {.FileKey = _fileKey, .Key = key, .Value = value, .LastModified = lastModified}
    Resources(r.Key) = r
   End If
  End Sub

  Public Sub New(fileKey As String, resourceFile As String)
   MyBase.New()

   _fileKey = fileKey
   _filePath = resourceFile

   If IO.File.Exists(resourceFile) Then
    Me.Load(resourceFile)
    For Each x As XmlNode In Me.DocumentElement.SelectNodes("/root/data")
     Dim r As New Resource
     r.Key = x.Attributes("name").InnerText
     'r.Value = x.SelectSingleNode("value").InnerXml
     r.Value = x.SelectSingleNode("value").InnerText
     If x.Attributes("lastModified") IsNot Nothing Then
      r.LastModified = CDate(x.Attributes("lastModified").InnerText)
     End If
     r.FileKey = _fileKey
     Resources(r.Key) = r
    Next
   Else
    Using s As IO.Stream = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("DotNetNuke.Translator.TextFile1.txt")
     Me.Load(s)
    End Using
   End If

  End Sub

  Public Sub Regenerate(minimal As Boolean)

   Using s As IO.Stream = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("DotNetNuke.Translator.TextFile1.txt")
    Me.Load(s)
   End Using
   'Me.DocumentElement.SelectSingleNode("/root").InnerXml = ""
   If Copyright <> "" Then
    Me.DocumentElement.SelectSingleNode("/root").AppendChild(Me.CreateComment(Copyright))
   End If
   For Each k As KeyValuePair(Of String, Resource) In Resources.OrderBy(Function(u) u.Value.Key)
    If Not String.IsNullOrEmpty(k.Value.Value) Then
     Dim x As XmlNode = Me.CreateElement("data")
     Me.DocumentElement.SelectSingleNode("/root").AppendChild(x)
     Dim a As XmlAttribute = Me.CreateAttribute("name")
     x.Attributes.Append(a)
     a.InnerText = k.Key
     If Not minimal AndAlso k.Value.LastModified <> DateTime.MinValue Then
      Dim lm As XmlAttribute = Me.CreateAttribute("lastModified")
      x.Attributes.Append(lm)
      lm.InnerText = k.Value.LastModified.ToUniversalTime.ToString("u").Substring(0, 19)
     End If
     Dim v As XmlNode = Me.CreateElement("value")
     x.AppendChild(v)
     'v.InnerXml = k.Value.Value
     v.InnerText = k.Value.Value
     Dim space As XmlAttribute = Me.CreateAttribute("xml:space")
     v.Attributes.Append(space)
     space.InnerText = "preserve"
    End If
   Next

  End Sub

  Public Overloads Sub Save()
   If Not String.IsNullOrEmpty(_filePath) Then

    Regenerate(False)
    Save(_FilePath)

   End If
  End Sub

 End Class

 Public Class Resource
  Public Property FileKey As String = ""
  Public Property Key As String = ""
  Public Property Value As String = ""
  Public Property LastModified As DateTime = DateTime.MinValue
 End Class
End Namespace