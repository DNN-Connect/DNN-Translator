Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Media

Namespace ViewModel
 Public Class QatItem
  Public Sub New()
  End Sub

  Public Sub New(instance__1 As Object, isSplitHeader__2 As Boolean)
   Instance = instance__1
   IsSplitHeader = isSplitHeader__2
  End Sub

  Public Property TabIndex() As Integer
   Get
    Return m_TabIndex
   End Get
   Set(value As Integer)
    m_TabIndex = Value
   End Set
  End Property
  Private m_TabIndex As Integer
  Public Property GroupIndex() As Integer
   Get
    Return m_GroupIndex
   End Get
   Set(value As Integer)
    m_GroupIndex = Value
   End Set
  End Property
  Private m_GroupIndex As Integer

  Public Property ControlIndices() As Int32Collection
   Get
    If _controlIndices Is Nothing Then
     _controlIndices = New Int32Collection()
    End If
    Return _controlIndices
   End Get
   Set(value As Int32Collection)
    _controlIndices = value
   End Set
  End Property
  Private _controlIndices As Int32Collection

  <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property Instance() As Object
   Get
    Return m_Instance
   End Get
   Set(value As Object)
    m_Instance = Value
   End Set
  End Property
  Private m_Instance As Object

  <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property IsSplitHeader() As Boolean
   Get
    Return m_IsSplitHeader
   End Get
   Set(value As Boolean)
    m_IsSplitHeader = Value
   End Set
  End Property
  Private m_IsSplitHeader As Boolean
 End Class
End Namespace
