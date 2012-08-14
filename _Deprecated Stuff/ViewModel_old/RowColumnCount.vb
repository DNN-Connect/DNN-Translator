Namespace ViewModel
 ''' <summary>
 ''' DataModel for Tables Gallery
 ''' </summary>
 Public Class RowColumnCount
  Public Property RowCount() As Integer
   Get
    Return m_RowCount
   End Get
   Set(value As Integer)
    m_RowCount = Value
   End Set
  End Property
  Private m_RowCount As Integer
  Public Property ColumnCount() As Integer
   Get
    Return m_ColumnCount
   End Get
   Set(value As Integer)
    m_ColumnCount = Value
   End Set
  End Property
  Private m_ColumnCount As Integer
 End Class
End Namespace
