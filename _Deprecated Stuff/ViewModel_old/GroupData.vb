Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Windows.Input

Namespace ViewModel
 Public Class GroupData
  Inherits ControlData
  Public Sub New(header As String)
   Label = header
  End Sub

  <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
  Public ReadOnly Property ControlDataCollection() As ObservableCollection(Of ControlData)
   Get
    If _controlDataCollection Is Nothing Then
     _controlDataCollection = New ObservableCollection(Of ControlData)()

     Dim smallImage As New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative)
     Dim largeImage As New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative)

     For i As Integer = 0 To ViewModelData.ButtonCount - 1
      _controlDataCollection.Add(New ButtonData() With {
       .Label = "Button " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
     For i As Integer = 0 To ViewModelData.ToggleButtonCount - 1
      _controlDataCollection.Add(New ToggleButtonData() With {
       .Label = "ToggleButton " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
     For i As Integer = 0 To ViewModelData.RadioButtonCount - 1
      _controlDataCollection.Add(New RadioButtonData() With {
       .Label = "RadioButton " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
     For i As Integer = 0 To ViewModelData.CheckBoxCount - 1
      _controlDataCollection.Add(New CheckBoxData() With {
       .Label = "CheckBox " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
     For i As Integer = 0 To ViewModelData.TextBoxCount - 1
      _controlDataCollection.Add(New TextBoxData() With {
       .Label = "TextBox " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
     For i As Integer = 0 To ViewModelData.MenuButtonCount - 1
      _controlDataCollection.Add(New MenuButtonData() With {
       .Label = "MenuButton " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
     For i As Integer = 0 To ViewModelData.SplitButtonCount - 1
      _controlDataCollection.Add(New SplitButtonData() With {
       .Label = "SplitButton " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand,
       .IsCheckable = True
      })
     Next
     For i As Integer = 0 To ViewModelData.ComboBoxCount - 1
      _controlDataCollection.Add(New ComboBoxData() With {
       .Label = "ComboBox " & i,
       .SmallImage = smallImage,
       .LargeImage = largeImage,
       .ToolTipTitle = "ToolTip Title",
       .ToolTipDescription = "ToolTip Description",
       .ToolTipImage = smallImage,
       .Command = ViewModelData.DefaultCommand
      })
     Next
    End If
    Return _controlDataCollection
   End Get
  End Property
  Private _controlDataCollection As ObservableCollection(Of ControlData)

 End Class
End Namespace
