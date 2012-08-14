Imports System.Windows.Input
Imports DotNetNuke.Translator.Common

Namespace ViewModel
 Public NotInheritable Class ViewModelData
  Private Sub New()
  End Sub
  Friend Const TabCount As Integer = 4
  Friend Const ContextualTabGroupCount As Integer = 2
  Friend Const GroupCount As Integer = 3
  Friend Const ControlCount As Integer = 5
  Friend Const ButtonCount As Integer = 1
  Friend Const ToggleButtonCount As Integer = 1
  Friend Const RadioButtonCount As Integer = 1
  Friend Const CheckBoxCount As Integer = 1
  Friend Const TextBoxCount As Integer = 1
  Friend Const MenuButtonCount As Integer = 1
  Friend Const MenuItemCount As Integer = 2
  Friend Const SplitButtonCount As Integer = 1
  Friend Const SplitMenuItemCount As Integer = 2
  Friend Const GalleryCount As Integer = 1
  Friend Const GalleryCategoryCount As Integer = 3
  Friend Const GalleryItemCount As Integer = 10
  Friend Const MenuItemNestingCount As Integer = 2
  Friend Const ComboBoxCount As Integer = 1

  Public Shared ReadOnly Property RibbonData() As RibbonData
   Get
    If _data Is Nothing Then
     _data = New RibbonData()
    End If
    Return _data
   End Get
  End Property

  Public Shared ReadOnly Property DefaultCommand() As ICommand
   Get
    If _defaultCommand Is Nothing Then
     _defaultCommand = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute)
    End If
    Return _defaultCommand
   End Get
  End Property

  Private Shared Sub DefaultExecuted()
  End Sub

  Private Shared Function DefaultCanExecute() As Boolean
   Return True
  End Function

  <ThreadStatic()> _
  Private Shared _data As RibbonData
  Private Shared _defaultCommand As ICommand
 End Class
End Namespace
