Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Namespace ViewModel
 Public Class ApplicationSplitMenuItemData
  Inherits SplitMenuItemData
  Public Sub New()
   Me.New(False)
  End Sub

  Public Sub New(isApplicationMenu As Boolean)
   MyBase.New(isApplicationMenu)
  End Sub
 End Class
End Namespace
