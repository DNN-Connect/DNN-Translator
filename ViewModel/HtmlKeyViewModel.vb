Imports System.ComponentModel
Imports DotNetNuke.Translator.Common

Namespace ViewModel
 Class HtmlKeyViewModel
  Inherits WorkspaceViewModel

  Public Property MainWindow As MainWindowViewModel = Nothing
  Public Property ResourceKey As ResourceKeyViewModel
  Public Property HasChanges As Boolean = False
  Public Property TargetLocale() As CultureInfo

  Public Sub New(mainWindow As MainWindowViewModel, resourceKey As ResourceKeyViewModel, targetLocale As CultureInfo)
   Me.MainWindow = mainWindow
   Me.ResourceKey = resourceKey.Clone
   Me.TargetLocale = targetLocale
   Me.ID = resourceKey.ID
   Me.DisplayName = resourceKey.Key

   AddHandler Me.ResourceKey.PropertyChanged, AddressOf ResourceChanged
  End Sub

  Protected Overrides Sub OnRequestClose()
   If HasChanges Then
    Dim messageBoxText As String = "Do you want to save changes?"
    Dim caption As String = "Word Processor"
    Dim button As MessageBoxButton = MessageBoxButton.YesNoCancel
    Dim icon As MessageBoxImage = MessageBoxImage.Warning
    Dim result As MessageBoxResult = MessageBox.Show(messageBoxText, caption, button, icon)
    Select Case result
     Case MessageBoxResult.Yes
      For Each ws As WorkspaceViewModel In MainWindow.Workspaces
       If TypeOf (ws) Is ResourceCollectionViewModel Then
        Dim rcvm As ResourceCollectionViewModel = CType(ws, ResourceCollectionViewModel)
        For Each k As ResourceKeyViewModel In rcvm.ResourceKeys
         If k.ID = Me.ID Then
          k.TargetValue = Me.ResourceKey.TargetValue
         End If
        Next
       End If
      Next
      Dim f As New ResourceFile(ResourceKey.ResourceFile, Common.Globals.GetResourceFilePath(MainWindow.ProjectSettings.Location, ResourceKey.ResourceFile, TargetLocale.Name))
      f.SetResourceValue(ResourceKey.Key, ResourceKey.TargetValue)
      f.Save()
     Case MessageBoxResult.No
      ' continue closing
     Case MessageBoxResult.Cancel
      Exit Sub
    End Select
   End If
   CloseMe()
  End Sub

#Region " Resource Changed Handler "
  Friend Overridable Sub ResourceChanged(sender As Object, e As PropertyChangedEventArgs)
   Select Case e.PropertyName
    Case "Selected"
    Case "TargetValue"
     HasChanges = True
   End Select
  End Sub
#End Region

 End Class
End Namespace
