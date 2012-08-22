Imports System.Security
Imports System.Text.RegularExpressions

Namespace Common
 Public Class Globals
  Public Shared Function XmlToFormattedByteArray(doc As Xml.XmlDocument) As Byte()
   Using w As New IO.MemoryStream
    Using xw As New Xml.XmlTextWriter(w, Text.Encoding.UTF8)
     xw.Formatting = Xml.Formatting.Indented
     doc.WriteContentTo(xw)
    End Using
    Return w.ToArray()
   End Using
  End Function

  Public Shared Function Clone(Of T)(source As T) As T
   If Not GetType(T).IsSerializable Then
    Throw New ArgumentException("The type must be serializable.", "source")
   End If
   If [Object].ReferenceEquals(source, Nothing) Then
    Return Nothing
   End If
   Dim formatter As Runtime.Serialization.IFormatter = New Runtime.Serialization.Formatters.Binary.BinaryFormatter()
   Using stream As IO.Stream = New IO.MemoryStream()
    formatter.Serialize(stream, source)
    stream.Seek(0, IO.SeekOrigin.Begin)
    Return DirectCast(formatter.Deserialize(stream), T)
   End Using
  End Function

  Public Shared Function AllLocales() As List(Of CultureInfo)
   Return AllLocales(CultureTypes.AllCultures)
  End Function
  Public Shared Function AllLocales(types As Globalization.CultureTypes) As List(Of CultureInfo)
   Dim res As List(Of CultureInfo) = CultureInfo.GetCultures(types).ToList()
   res.Sort(Function(l1, l2) l1.EnglishName.CompareTo(l2.EnglishName))
   Return res
  End Function

  Public Shared Function GetResourceFilePath(installationPath As String, resourceFile As String, locale As String) As String
   resourceFile = Text.RegularExpressions.Regex.Replace(resourceFile, "(\.en\-US)?\.resx$", "")
   Return String.Format("{0}{1}.{2}.resx", installationPath, resourceFile, locale)
  End Function
  Public Shared Function GetResourceFileKey(installationPath As String, resourceFile As String) As String
   resourceFile = resourceFile.Substring(installationPath.Length)
   resourceFile = resourceFile.Trim("\"c)
   resourceFile = Text.RegularExpressions.Regex.Replace(resourceFile, "(\.\w\w\-\w\w)?\.resx$", "")
   Return resourceFile
  End Function

  Public Shared Function GetTranslation(installationPath As String, resourceFile As String, resourceKey As String, locale As String) As String
   Dim resFile As String = GetResourceFilePath(installationPath, resourceFile, locale)
   If IO.File.Exists(resFile) Then
    Dim rf As New ResourceFile(resourceFile, resFile)
    If rf.Resources.ContainsKey(resourceKey) Then
     Return rf.Resources(resourceKey).Value
    Else
     Return ""
    End If
   End If
   Return ""
  End Function

  Public Shared Function NormalizeVersion(version As String) As String
   Dim m As Match = Regex.Match(version, "(\d+)\.(\d+)\.(\d+)(\.\d+)?")
   If m.Success Then
    Return String.Format("{0:00}.{1:00}.{2:00}", CInt(m.Groups(1).Value), CInt(m.Groups(2).Value), CInt(m.Groups(3).Value))
   Else
    Return version
   End If
  End Function

#Region " Encryption "
  Shared entropy As Byte() = System.Text.Encoding.Unicode.GetBytes("Not A Password")

  Public Shared Function EncryptString(input As System.Security.SecureString) As String
   If input Is Nothing Then Return ""
   Dim encryptedData As Byte() = System.Security.Cryptography.ProtectedData.Protect(System.Text.Encoding.Unicode.GetBytes(ToInsecureString(input)), entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser)
   Return Convert.ToBase64String(encryptedData)
  End Function

  Public Shared Function DecryptString(encryptedData As String) As SecureString
   Try
    Dim decryptedData As Byte() = System.Security.Cryptography.ProtectedData.Unprotect(Convert.FromBase64String(encryptedData), entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser)
    Return ToSecureString(System.Text.Encoding.Unicode.GetString(decryptedData))
   Catch
    Return New SecureString()
   End Try
  End Function

  Public Shared Function ToSecureString(input As String) As SecureString
   Dim secure As New SecureString()
   For Each c As Char In input
    secure.AppendChar(c)
   Next
   secure.MakeReadOnly()
   Return secure
  End Function

  Public Shared Function ToInsecureString(input As SecureString) As String
   If input Is Nothing Then Return ""
   Dim returnValue As String = String.Empty
   Dim ptr As IntPtr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input)
   Try
    returnValue = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr)
   Finally
    System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr)
   End Try
   Return returnValue
  End Function
#End Region

 End Class
End Namespace
