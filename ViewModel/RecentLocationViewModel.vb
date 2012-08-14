Namespace ViewModel
 Public Class RecentLocationViewModel
  Inherits CommandViewModel

  Private Property RecentLocation As Data.RecentLocation

  Public Sub New(location As Data.RecentLocation, command As ICommand)
   MyBase.New("RecentLocation", command)
   Me.RecentLocation = location
  End Sub

  Public ReadOnly Property Location As String
   Get
    Return RecentLocation.Location
   End Get
  End Property

  Public ReadOnly Property LocationShort As String
   Get
    Return RecentLocation.LocationShort
   End Get
  End Property

  Public ReadOnly Property Index As Integer
   Get
    Return RecentLocation.Index
   End Get
  End Property
 End Class
End Namespace
