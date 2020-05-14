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

  Public Shared Function GetLocalizedFilePath(resourceFile As String, locale As String) As String
   If resourceFile.ToLower.EndsWith("template.en-us.resx") Then
    Return resourceFile.Substring(0, resourceFile.Length - 11) & "." & locale & ".resx"
   Else
    Return resourceFile.Substring(0, resourceFile.Length - 5) & "." & locale & ".resx"
   End If
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

  Public Shared Function TranslatorDocFolder() As String
   Dim res As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "DNNTranslator")
   If Not IO.Directory.Exists(res) Then IO.Directory.CreateDirectory(res)
   Return res
  End Function

#Region " UI "
  Public Shared Function FindChild(Of T As DependencyObject)(parent As DependencyObject, childName As String) As T
   ' Confirm parent and childName are valid. 
   If parent Is Nothing Then
    Return Nothing
   End If

   Dim foundChild As T = Nothing

   Dim childrenCount As Integer = VisualTreeHelper.GetChildrenCount(parent)
   For i As Integer = 0 To childrenCount - 1
    Dim child = VisualTreeHelper.GetChild(parent, i)
    ' If the child is not of the request child type child
    Dim childType As T = TryCast(child, T)
    If childType Is Nothing Then
     ' recursively drill down the tree
     foundChild = FindChild(Of T)(child, childName)

     ' If the child is found, break so we do not overwrite the found child. 
     If foundChild IsNot Nothing Then
      Exit For
     End If
    ElseIf Not String.IsNullOrEmpty(childName) Then
     Dim frameworkElement = TryCast(child, FrameworkElement)
     ' If the child's name is set for search
     If frameworkElement IsNot Nothing AndAlso frameworkElement.Name = childName Then
      ' if the child's name is of the request name
      foundChild = DirectCast(child, T)
      Exit For
     End If
    Else
     ' child element found.
     foundChild = DirectCast(child, T)
     Exit For
    End If
   Next

   Return foundChild
  End Function
#End Region

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

#Region " Html/XML Encoding "
    ''' <summary>
    ''' Not used for now. May come in handy later.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getEncodingDictionary() As String(,)

   Dim s As String = "&Aacute;,193,C1,Á|"
   s = s & "&aacute;,225,E1,á|"
   s = s & "&Acirc;,194,C2,Â|"
   s = s & "&acirc;,226,E2,â|"
   s = s & "&acute;,180,B4,´|"
   s = s & "&AElig;,198,C6,Æ|"
   s = s & "&aelig;,230,E6,æ|"
   s = s & "&Agrave;,192,C0,À|"
   s = s & "&agrave;,224,E0,à|"
   s = s & "&Aring;,197,C5,Å|"
   s = s & "&aring;,229,E5,å|"
   s = s & "&Atilde;,195,C3,Ã|"
   s = s & "&atilde;,227,E3,ã|"
   s = s & "&Auml;,196,C4,Ä|"
   s = s & "&auml;,228,E4,ä|"
   s = s & "&bdquo;,132,84,„|"
   s = s & "&brvbar;,166,A6,¦|"
   s = s & "&bull;,149,95,•|"
   s = s & "&Ccedil;,199,C7,Ç|"
   s = s & "&ccedil;,231,E7,ç|"
   s = s & "&cedil;,184,B8,¸|"
   s = s & "&cent;,162,A2,¢|"
   s = s & "&circ;,136,88,ˆ|"
   s = s & "&copy;,169,A9,©|"
   s = s & "&curren;,164,A4,¤|"
   s = s & "&dagger;,134,86,†|"
   s = s & "&Dagger;,135,87,‡|"
   s = s & "&deg;,176,B0,°|"
   s = s & "&divide;,247,F7,÷|"
   s = s & "&Eacute;,201,C9,É|"
   s = s & "&eacute;,233,E9,é|"
   s = s & "&Ecirc;,202,CA,Ê|"
   s = s & "&ecirc;,234,EA,ê|"
   s = s & "&Egrave;,200,C8,È|"
   s = s & "&egrave;,232,E8,è|"
   s = s & "&ETH;,208,D0,Ð|"
   s = s & "&eth;,240,F0,ð|"
   s = s & "&Euml;,203,CB,Ë|"
   s = s & "&euml;,235,EB,ë|"
   s = s & "&euro;,128,80,€|"
   s = s & "&fnof;,131,83,ƒ|"
   s = s & "&frac12;,189,BD,½|"
   s = s & "&frac14;,188,BC,¼|"
   s = s & "&frac34;,190,BE,¾|"
   s = s & "&hellip;,133,85,…|"
   s = s & "&Iacute;,205,CD,Í|"
   s = s & "&iacute;,237,ED,í|"
   s = s & "&Icirc;,206,CE,Î|"
   s = s & "&icirc;,238,EE,î|"
   s = s & "&iexcl;,161,A1,¡|"
   s = s & "&Igrave;,204,CC,Ì|"
   s = s & "&igrave;,236,EC,ì|"
   s = s & "&iquest;,191,BF,¿|"
   s = s & "&Iuml;,207,CF,Ï|"
   s = s & "&iuml;,239,EF,ï|"
   s = s & "&laquo;,171,AB,«|"
   s = s & "&ldquo;,147,93,""|"
   s = s & "&lsaquo;,139,8B,‹|"
   s = s & "&lsquo;,145,91,‘|"
   s = s & "&macr;,175,AF,¯|"
   s = s & "&mdash;,45,97,—|"
   s = s & "&micro;,181,B5,µ|"
   s = s & "&middot;,183,B7,•|"
   s = s & "&nbsp;,160,A0, |"
   s = s & "&ndash;,45,96,–|"
   s = s & "&not;,172,AC,¬|"
   s = s & "&Ntilde;,209,D1,Ñ|"
   s = s & "&ntilde;,241,F1,ñ|"
   s = s & "&Oacute;,211,D3,Ó|"
   s = s & "&oacute;,243,F3,ó|"
   s = s & "&Ocirc;,212,D4,Ô|"
   s = s & "&ocirc;,244,F4,ô|"
   s = s & "&OElig;,140,8C,Œ|"
   s = s & "&oelig;,156,9C,œ|"
   s = s & "&Ograve;,210,D2,Ò|"
   s = s & "&ograve;,242,F2,ò|"
   s = s & "&ordf;,170,AA,ª|"
   s = s & "&ordm;,186,BA,º|"
   s = s & "&Oslash;,216,D8,Ø|"
   s = s & "&oslash;,248,F8,ø|"
   s = s & "&Otilde;,213,D5,Õ|"
   s = s & "&otilde;,245,F5,õ|"
   s = s & "&Ouml;,214,D6,Ö|"
   s = s & "&ouml;,246,F6,ö|"
   s = s & "&para;,182,B6,¶|"
   s = s & "&permil;,137,89,‰|"
   s = s & "&plusmn;,177,B1,±|"
   s = s & "&pound;,163,A3,£|"
   s = s & "&raquo;,187,BB,»|"
   s = s & "&rdquo;,148,94,""|"
   s = s & "&reg;,174,AE,®|"
   s = s & "&rsaquo;,155,9B,›|"
   s = s & "&rsquo;,146,',’|"
   s = s & "&sbquo;,130,82,‚|"
   s = s & "&Scaron;,138,8A,Š|"
   s = s & "&scaron;,154,9A,š|"
   s = s & "&sect;,167,A7,§|"
   s = s & "&shy;,173,AD,¬|"
   s = s & "&sup1;,185,B9,¹|"
   s = s & "&sup2;,178,B2,²|"
   s = s & "&sup3;,179,B3,³|"
   s = s & "&szlig;,223,DF,ß|"
   s = s & "&THORN;,222,DE,Þ|"
   s = s & "&thorn;,254,FE,þ|"
   s = s & "&tilde;,152,98,˜|"
   s = s & "&times;,215,D7,×|"
   s = s & "&trade;,153,99,™|"
   s = s & "&Uacute;,218,DA,Ú|"
   s = s & "&uacute;,250,FA,ú|"
   s = s & "&Ucirc;,219,DB,Û|"
   s = s & "&ucirc;,251,FB,û|"
   s = s & "&Ugrave;,217,D9,Ù|"
   s = s & "&ugrave;,249,F9,ù|"
   s = s & "&uml;,168,A8,¨|"
   s = s & "&Uuml;,220,DC,Ü|"
   s = s & "&uuml;,252,FC,ü|"
   s = s & "&Yacute;,221,DD,Ý|"
   s = s & "&yacute;,253,FD,ý|"
   s = s & "&yen;,165,A5,¥|"
   s = s & "&yuml;,159,9F,Ÿ|"
   s = s & "&yuml;,255,FF,ÿ"
   s = s & "&#39;,',','"
   s = s & "<ul>,,,"
   s = s & "<li>, - , - , - "
   s = s & "</li>,,,"
   s = s & "</ul>,,,"

   Dim records() As String = s.Split("|"c)
   Dim result(UBound(records), 3) As String
   For i As Integer = 0 To UBound(records) - 1
    Dim rec() As String = records(i).Split(","c)
    result(i, 0) = rec(0)
    result(i, 1) = rec(1)
    result(i, 2) = rec(2)
    result(i, 3) = rec(3)
   Next

   Return result

  End Function
#End Region

 End Class
End Namespace
