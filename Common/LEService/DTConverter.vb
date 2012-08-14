Namespace Common.LEService
 Public Class DTConverter
  Inherits Newtonsoft.Json.Converters.DateTimeConverterBase

  Private Shared ReadOnly Epoch As New DateTime(1970, 1, 1)

  Public Overrides Function ReadJson(reader As Newtonsoft.Json.JsonReader, objectType As System.Type, existingValue As Object, serializer As Newtonsoft.Json.JsonSerializer) As Object
   Return Date.MinValue
  End Function

  Public Overrides Sub WriteJson(writer As Newtonsoft.Json.JsonWriter, value As Object, serializer As Newtonsoft.Json.JsonSerializer)
   Dim v As DateTime = CType(value, DateTime)
   writer.WriteValue(String.Format("/Date({0:0})/", (v - Epoch).TotalMilliseconds))
  End Sub
 End Class
End Namespace
