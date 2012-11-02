Imports DotNetNuke.Translator.ViewModel
Imports DotNetNuke.Translator.Common

Namespace View
 Partial Class ResourceFileView
  Inherits System.Windows.Controls.UserControl

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
 End Class
End Namespace