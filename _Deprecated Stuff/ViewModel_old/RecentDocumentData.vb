Imports System.ComponentModel

Namespace ViewModel
 Public Class RecentDocumentData
  Inherits ToggleButtonData
  Public Property Index() As Integer
   Get
    Return _index
   End Get

   Set(value As Integer)
    If _index <> value Then
     _index = value
     OnPropertyChanged(New PropertyChangedEventArgs("Index"))
    End If
   End Set
  End Property
  Private _index As Integer

  Public ReadOnly Property LabelShort As String
   Get
    If Label.Length > 50 Then
     Return Label.Substring(0, 3) & "..." & Label.Substring(Label.Length - 40, 40)
    Else
     Return Label
    End If
   End Get
  End Property
 End Class
End Namespace
