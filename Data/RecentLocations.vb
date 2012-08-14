Imports System
Imports System.IO
Imports System.IO.IsolatedStorage

Namespace Data
 Public Class RecentLocations

  Private Const RecentLocationsFilename As String = "RecentLocations.txt"
  Public Property RecentLocations As New SortedList(Of Integer, RecentLocation)

  Public Sub New()
   Dim isoStore As IsolatedStorageFile = IsolatedStorageFile.GetStore(IsolatedStorageScope.User Or IsolatedStorageScope.Assembly Or IsolatedStorageScope.Domain, Nothing, Nothing)
   Dim recentFiles As String() = isoStore.GetFileNames(RecentLocationsFilename)
   Dim added As New List(Of String)
   If recentFiles.Length > 0 Then
    Dim i As Integer = 1
    Using strIn As New IO.StreamReader(New IsolatedStorageFileStream(RecentLocationsFilename, FileMode.Open, isoStore))
     Do While Not strIn.EndOfStream
      Dim s As String = strIn.ReadLine
      If Not String.IsNullOrEmpty(s) AndAlso IO.File.Exists(s) AndAlso Not added.Contains(s) Then
       RecentLocations.Add(i, New RecentLocation(i, s))
       added.Add(s)
       i += 1
      End If
     Loop
    End Using
   End If
  End Sub

  Public Sub Remove(location As String)
   Dim tmpLocations As New SortedList(Of Integer, RecentLocation)
   Dim i As Integer = 1
   For Each kvp As KeyValuePair(Of Integer, RecentLocation) In RecentLocations
    If kvp.Value.Location <> location Then
     tmpLocations.Add(i, New RecentLocation(i, kvp.Value.Location))
     i += 1
    End If
   Next
   RecentLocations = tmpLocations
  End Sub

  Public Sub Push(location As String)
   Dim tmpLocations As New SortedList(Of Integer, RecentLocation)
   tmpLocations.Add(1, New RecentLocation(1, location))
   Dim i As Integer = 2
   For Each kvp As KeyValuePair(Of Integer, RecentLocation) In RecentLocations
    If kvp.Value.Location <> location Then
     tmpLocations.Add(i, New RecentLocation(i, kvp.Value.Location))
     i += 1
    End If
   Next
   RecentLocations = tmpLocations
  End Sub

  Public Sub Save()
   Dim isoStore As IsolatedStorageFile = IsolatedStorageFile.GetStore(IsolatedStorageScope.User Or IsolatedStorageScope.Assembly Or IsolatedStorageScope.Domain, Nothing, Nothing)
   Using strOut As New IO.StreamWriter(New IsolatedStorageFileStream(RecentLocationsFilename, FileMode.OpenOrCreate, isoStore))
    For Each kvp As KeyValuePair(Of Integer, RecentLocation) In RecentLocations
     strOut.WriteLine(kvp.Value.Location)
     If kvp.Key > 10 Then Exit For ' maximum of 10 recent locations
    Next
   End Using
  End Sub

 End Class
End Namespace
