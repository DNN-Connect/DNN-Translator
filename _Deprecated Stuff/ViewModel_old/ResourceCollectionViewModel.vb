Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Input

Namespace ViewModel
 Public Class ResourceCollectionViewModel

#Region "Data"

  ReadOnly _topLevel As ReadOnlyCollection(Of ResourceFileViewModel)
  ReadOnly _rootPerson As ResourceFileViewModel
  ReadOnly _searchCommand As ICommand

  Private _matchingPeopleEnumerator As IEnumerator(Of ResourceFileViewModel)
  Private _searchText As String = [String].Empty

#End Region

#Region "Constructor"

  Public Sub New(root As Person)
   _rootPerson = New ResourceFileViewModel(root)

   _topLevel = New ReadOnlyCollection(Of ResourceFileViewModel)(New ResourceFileViewModel() {_rootPerson})

   '_searchCommand = New SearchFamilyTreeCommand(Me)
  End Sub

#End Region

#Region "Properties"

#Region "FirstGeneration"

  ''' <summary>
  ''' Returns a read-only collection containing the first person 
  ''' in the family tree, to which the TreeView can bind.
  ''' </summary>
  Public ReadOnly Property TopLevel() As ReadOnlyCollection(Of ResourceFileViewModel)
   Get
    Return _topLevel
   End Get
  End Property

#End Region

  '#Region "SearchCommand"

  '  ''' <summary>
  '  ''' Returns the command used to execute a search in the family tree.
  '  ''' </summary>
  '  Public ReadOnly Property SearchCommand() As ICommand
  '   Get
  '    Return _searchCommand
  '   End Get
  '  End Property

  '  Private Class SearchFamilyTreeCommand
  '   Implements ICommand
  '   ReadOnly _familyTree As FamilyTreeViewModel

  '   Public Sub New(familyTree As FamilyTreeViewModel)
  '    _familyTree = familyTree
  '   End Sub

  '   Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
  '    Return True
  '   End Function

  '   Private Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
  '    AddHandler(ByVal value As EventHandler)
  '     ' I intentionally left these empty because
  '     ' this command never raises the event, and
  '     ' not using the WeakEvent pattern here can
  '     ' cause memory leaks.  WeakEvent pattern is
  '     ' not simple to implement, so why bother.
  '    End AddHandler
  '    RemoveHandler(ByVal value As EventHandler)
  '    End RemoveHandler
  '   End Event

  '   Public Sub Execute(parameter As Object) Implements ICommand.Execute
  '    _familyTree.PerformSearch()
  '   End Sub
  '  End Class

  '#End Region

  '#Region "SearchText"

  '  ''' <summary>
  '  ''' Gets/sets a fragment of the name to search for.
  '  ''' </summary>
  '  Public Property SearchText() As String
  '   Get
  '    Return _searchText
  '   End Get
  '   Set(value As String)
  '    If value = _searchText Then
  '     Return
  '    End If

  '    _searchText = value

  '    _matchingPeopleEnumerator = Nothing
  '   End Set
  '  End Property

  '#End Region

#End Region

 End Class
End Namespace
