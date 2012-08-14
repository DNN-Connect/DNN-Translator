Namespace Common.LEService
 Public Class TextInfo
  Public Property DeprecatedIn() As String
  Public Property FilePath() As String
  Public Property ObjectId() As Int32
  Public Property OriginalValue() As String
  Public Property TextId() As Int32
  Public Property TextKey() As String
  Public Property Version() As String
  Public Property TextValue() As String
  Public Property Locale() As String
  Public Property Translation As String = ""
  Public Property LastModified As DateTime = DateTime.MinValue
  Public Property LastModifiedUserId As Integer = -1
 End Class
End Namespace
