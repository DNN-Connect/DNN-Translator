Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel
Imports System.Windows.Input

Public Class ControlData
 Implements INotifyPropertyChanged

 Public Property Label() As String
  Get
   Return _label
  End Get

  Set(value As String)
   If _label <> value Then
    _label = value
    OnPropertyChanged(New PropertyChangedEventArgs("Label"))
   End If
  End Set
 End Property
 Private _label As String

 Public Property LargeImage() As Uri
  Get
   Return _largeImage
  End Get

  Set(value As Uri)
   If _largeImage IsNot value Then
    _largeImage = value
    OnPropertyChanged(New PropertyChangedEventArgs("LargeImage"))
   End If
  End Set
 End Property
 Private _largeImage As Uri

 Public Property SmallImage() As Uri
  Get
   Return _smallImage
  End Get

  Set(value As Uri)
   If _smallImage IsNot value Then
    _smallImage = value
    OnPropertyChanged(New PropertyChangedEventArgs("SmallImage"))
   End If
  End Set
 End Property
 Private _smallImage As Uri

 Public Property ToolTipTitle() As String
  Get
   Return _toolTipTitle
  End Get

  Set(value As String)
   If _toolTipTitle <> value Then
    _toolTipTitle = value
    OnPropertyChanged(New PropertyChangedEventArgs("ToolTipTitle"))
   End If
  End Set
 End Property
 Private _toolTipTitle As String

 Public Property ToolTipDescription() As String
  Get
   Return _toolTipDescription
  End Get

  Set(value As String)
   If _toolTipDescription <> value Then
    _toolTipDescription = value
    OnPropertyChanged(New PropertyChangedEventArgs("ToolTipDescription"))
   End If
  End Set
 End Property
 Private _toolTipDescription As String

 Public Property ToolTipImage() As Uri
  Get
   Return _toolTipImage
  End Get

  Set(value As Uri)
   If _toolTipImage IsNot value Then
    _toolTipImage = value
    OnPropertyChanged(New PropertyChangedEventArgs("ToolTipImage"))
   End If
  End Set
 End Property
 Private _toolTipImage As Uri

 Public Property ToolTipFooterTitle() As String
  Get
   Return _toolTipFooterTitle
  End Get

  Set(value As String)
   If _toolTipFooterTitle <> value Then
    _toolTipFooterTitle = value
    OnPropertyChanged(New PropertyChangedEventArgs("ToolTipFooterTitle"))
   End If
  End Set
 End Property
 Private _toolTipFooterTitle As String

 Public Property ToolTipFooterDescription() As String
  Get
   Return _toolTipFooterDescription
  End Get

  Set(value As String)
   If _toolTipFooterDescription <> value Then
    _toolTipFooterDescription = value
    OnPropertyChanged(New PropertyChangedEventArgs("ToolTipFooterDescription"))
   End If
  End Set
 End Property
 Private _toolTipFooterDescription As String

 Public Property ToolTipFooterImage() As Uri
  Get
   Return _toolTipFooterImage
  End Get

  Set(value As Uri)
   If _toolTipFooterImage IsNot value Then
    _toolTipFooterImage = value
    OnPropertyChanged(New PropertyChangedEventArgs("ToolTipFooterImage"))
   End If
  End Set
 End Property
 Private _toolTipFooterImage As Uri

 Public Property Command() As ICommand
  Get
   Return _command
  End Get

  Set(value As ICommand)
   If _command IsNot value Then
    _command = value
    OnPropertyChanged(New PropertyChangedEventArgs("Command"))
   End If
  End Set
 End Property
 Private _command As ICommand

 Public Property KeyTip() As String
  Get
   Return _keyTip
  End Get

  Set(value As String)
   If _keyTip <> value Then
    _keyTip = value
    OnPropertyChanged(New PropertyChangedEventArgs("KeyTip"))
   End If
  End Set
 End Property
 Private _keyTip As String

#Region "INotifyPropertyChanged Members"

 Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

 Protected Sub OnPropertyChanged(e As PropertyChangedEventArgs)
  RaiseEvent PropertyChanged(Me, e)
 End Sub

#End Region

End Class


