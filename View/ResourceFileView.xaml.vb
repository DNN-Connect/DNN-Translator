Imports DotNetNuke.Translator.ViewModel
Imports DotNetNuke.Translator.Common
Imports System.Windows.Controls.Primitives

Namespace View
 Partial Class ResourceFileView
  Inherits System.Windows.Controls.UserControl

  Private Sub ResourcesGrid_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles ResourcesGrid.Loaded
   For Each row As DataGridRow In ResourcesGrid.GetDataGridRows
    Dim targetText As TextBox = row.GetVisualChild(Of TextBox)("txtTarget")
    If targetText IsNot Nothing Then
     AddHandler targetText.KeyDown, AddressOf Me.TargetKeyDown
    End If
   Next
  End Sub

  Private Sub ResourcesGrid_SelectedCellsChanged(sender As Object, e As System.Windows.Controls.SelectedCellsChangedEventArgs) Handles ResourcesGrid.SelectedCellsChanged
   ' workaround for MVVM not cooperating setting the focus and selecting all text upon selection
   Try
    For Each row As DataGridRow In ResourcesGrid.GetDataGridRows
     If row.Item Is Nothing Then Exit Sub
     If CType(row.Item, ResourceKeyViewModel).Selected Then
      Dim tt As TextBox = Globals.FindChild(Of TextBox)(row, "txtTarget")
      tt.SelectAll()
      tt.Focus()
      Exit Sub
     End If
    Next
   Catch ex As Exception
    MsgBox(ex.Message, MsgBoxStyle.Critical)
   End Try
  End Sub

  Public Sub TargetKeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs)
   If e.Key.Equals(Key.Enter) AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
    Dim targetText As TextBox = CType(sender, TextBox)
    targetText.Text = targetText.Text.Insert(targetText.CaretIndex, vbCrLf)
   End If
  End Sub

 End Class
End Namespace