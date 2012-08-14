Namespace Model
 Public Class ResourceKey

#Region " Properties "
  Private _key As String
  Public ReadOnly Property Key As String
   Get
    Return _key
   End Get
  End Property

  Private _originalValue As String
  Public ReadOnly Property OriginalValue As String
   Get
    Return _originalValue
   End Get
  End Property

  Private _targetValue As String
  Public Property TargetValue() As String
   Get
    Return _targetValue
   End Get
   Set(ByVal value As String)
    _targetValue = value
   End Set
  End Property

  Private _compareValue As String
  Public Property CompareValue() As String
   Get
    Return _compareValue
   End Get
   Set(ByVal value As String)
    _compareValue = value
   End Set
  End Property

  Private _selected As Boolean = False
  Public Property Selected() As Boolean
   Get
    Return _selected
   End Get
   Set(ByVal value As Boolean)
    _selected = value
   End Set
  End Property
#End Region

#Region " Constructors "
  Public Sub New(key As String, originalValue As String)
   _key = key
   _originalValue = originalValue
  End Sub

  Public Sub New(key As String, originalValue As String, targetValue As String)
   _key = key
   _originalValue = originalValue
   _targetValue = targetValue
  End Sub

  Public Sub New(key As String, originalValue As String, targetValue As String, compareValue As String)
   _key = key
   _originalValue = originalValue
   _targetValue = targetValue
   _compareValue = compareValue
  End Sub
#End Region

#Region " Comparison "
  Public Overrides Function Equals(obj As Object) As Boolean
   If obj Is Nothing OrElse [GetType]() IsNot obj.[GetType]() Then
    Return False
   End If

   Dim other As ResourceKey = TryCast(obj, ResourceKey)

   If Me.Key <> other.Key Then
    Return False
   End If

   If Me.OriginalValue <> other.OriginalValue Then
    Return False
   End If

   Return True
  End Function

  Public Overrides Function GetHashCode() As Integer
   Return Key.GetHashCode() Xor OriginalValue.GetHashCode()
  End Function

  Public Shared Operator =(key1 As ResourceKey, key2 As ResourceKey) As Boolean
   If [Object].Equals(key1, Nothing) AndAlso [Object].Equals(key2, Nothing) Then
    Return True
   End If
   Return key1.Equals(key2)
  End Operator

  Public Shared Operator <>(key1 As ResourceKey, key2 As ResourceKey) As Boolean
   Return Not (key1 = key2)
  End Operator
#End Region

 End Class
End Namespace
