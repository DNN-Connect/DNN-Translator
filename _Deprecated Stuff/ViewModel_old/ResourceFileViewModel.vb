Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Input
Imports System.ComponentModel

Namespace ViewModel
 Public Class ResourceFileViewModel
  Implements INotifyPropertyChanged

#Region "Data"

  ReadOnly _children As ReadOnlyCollection(Of ResourceFileViewModel)
  ReadOnly _parent As ResourceFileViewModel
  ReadOnly _resourcefile As TranslatorData.ResourceFilesRow

  Private _isExpanded As Boolean
  Private _isSelected As Boolean

#End Region

#Region "Constructors"

  Public Sub New(resourcefile As TranslatorData.ResourceFilesRow)
   Me.New(resourcefile, Nothing)
  End Sub

  Private Sub New(resourcefile As TranslatorData.ResourceFilesRow, parent As ResourceFileViewModel)
   _resourcefile = resourcefile
   _parent = parent

			_children = New ReadOnlyCollection(Of ResourceFileViewModel)((From child In _resourcefile.Children New ResourceFileViewModel(child, Me)).ToList(Of ResourceFileViewModel)())
  End Sub

#End Region

#Region "resourcefile Properties"

  Public ReadOnly Property Children() As ReadOnlyCollection(Of ResourceFileViewModel)
   Get
    Return _children
   End Get
  End Property

  Public ReadOnly Property Name() As String
   Get
    Return _resourcefile.FilePath
   End Get
  End Property

#End Region

#Region "Presentation Members"

#Region "IsExpanded"

  ''' <summary>
  ''' Gets/sets whether the TreeViewItem 
  ''' associated with this object is expanded.
  ''' </summary>
  Public Property IsExpanded() As Boolean
   Get
    Return _isExpanded
   End Get
   Set(value As Boolean)
    If value <> _isExpanded Then
     _isExpanded = value
     Me.OnPropertyChanged("IsExpanded")
    End If

    ' Expand all the way up to the root.
    If _isExpanded AndAlso _parent IsNot Nothing Then
     _parent.IsExpanded = True
    End If
   End Set
  End Property

#End Region

#Region "IsSelected"

  ''' <summary>
  ''' Gets/sets whether the TreeViewItem 
  ''' associated with this object is selected.
  ''' </summary>
  Public Property IsSelected() As Boolean
   Get
    Return _isSelected
   End Get
   Set(value As Boolean)
    If value <> _isSelected Then
     _isSelected = value
     Me.OnPropertyChanged("IsSelected")
    End If
   End Set
  End Property

#End Region

#Region "NameContainsText"

  Public Function NameContainsText(text As String) As Boolean
   If [String].IsNullOrEmpty(text) OrElse [String].IsNullOrEmpty(Me.Name) Then
    Return False
   End If

   Return Me.Name.IndexOf(text, StringComparison.InvariantCultureIgnoreCase) > -1
  End Function

#End Region

#Region "Parent"

  Public ReadOnly Property Parent() As ResourceFileViewModel
   Get
    Return _parent
   End Get
  End Property

#End Region

#End Region

#Region "INotifyPropertyChanged Members"

  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Protected Overridable Sub OnPropertyChanged(propertyName As String)
   RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
  End Sub

#End Region


 End Class
End Namespace
