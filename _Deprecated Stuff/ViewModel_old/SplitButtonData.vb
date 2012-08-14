Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel

Namespace ViewModel
 Public Class SplitButtonData
  Inherits MenuButtonData
  Public Sub New()
   Me.New(False)
  End Sub

  Public Sub New(isApplicationMenu As Boolean)
   MyBase.New(isApplicationMenu)
  End Sub

  Public Property IsChecked() As Boolean
   Get
    Return _isChecked
   End Get

   Set(value As Boolean)
    If _isChecked <> value Then
     _isChecked = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsChecked"))
    End If
   End Set
  End Property
  Private _isChecked As Boolean

  Public Property IsCheckable() As Boolean
   Get
    Return _isCheckable
   End Get

   Set(value As Boolean)
    If _isCheckable <> value Then
     _isCheckable = value
     OnPropertyChanged(New PropertyChangedEventArgs("IsCheckable"))
    End If
   End Set
  End Property
  Private _isCheckable As Boolean

  Public ReadOnly Property DropDownButtonData() As ButtonData
   Get
    If _dropDownButtonData Is Nothing Then
     _dropDownButtonData = New ButtonData()
    End If

    Return _dropDownButtonData
   End Get
  End Property
  Private _dropDownButtonData As ButtonData
 End Class
End Namespace
