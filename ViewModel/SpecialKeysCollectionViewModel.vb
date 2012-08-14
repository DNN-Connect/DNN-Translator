Namespace ViewModel
 Public Class SpecialKeysCollectionViewModel
  Inherits ResourceCollectionViewModel

  Public Sub New(mainWindow As MainWindowViewModel, targetLocale As CultureInfo, collection As String)
   MyBase.New(mainWindow)

   Select Case collection.ToLower
    Case "missing"

    Case "ignored"

    Case "search"

   End Select

  End Sub

 End Class
End Namespace
