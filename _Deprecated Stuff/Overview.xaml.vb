Namespace Controls
 Public Class Overview

  Private Sub Overview_Initialized(sender As Object, e As System.EventArgs) Handles Me.Initialized
   AddHandler ViewModel.TranslatorModel.LocationChanged, AddressOf ChangedLocation
  End Sub

  Private Sub ChangedLocation()

  End Sub
 End Class
End Namespace
