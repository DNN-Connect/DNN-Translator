Imports System.Collections.ObjectModel

Namespace Common
 Public Class SortableObservableCollection(Of T)
  Inherits ObservableCollection(Of T)

  Public Sub New()
   MyBase.New()
  End Sub

  Public Sub New(list As List(Of T))
   MyBase.New(list)
  End Sub

  Public Sub New(collection As IEnumerable(Of T))
   MyBase.New(collection)
  End Sub

  Public Sub Sort(Of TKey)(keySelector As Func(Of T, TKey), direction As System.ComponentModel.ListSortDirection)
   Select Case direction
    Case System.ComponentModel.ListSortDirection.Ascending
     If True Then
      ApplySort(Items.OrderBy(keySelector))
      Exit Select
     End If
    Case System.ComponentModel.ListSortDirection.Descending
     If True Then
      ApplySort(Items.OrderByDescending(keySelector))
      Exit Select
     End If
   End Select
  End Sub

  Public Sub Sort(Of TKey)(keySelector As Func(Of T, TKey), comparer As IComparer(Of TKey))
   ApplySort(Items.OrderBy(keySelector, comparer))
  End Sub

  Private Sub ApplySort(sortedItems As IEnumerable(Of T))
   Dim sortedItemsList = sortedItems.ToList()

   For Each item As T In sortedItemsList
    Move(IndexOf(item), sortedItemsList.IndexOf(item))
   Next
  End Sub
 End Class
End Namespace
