Imports System.Collections.ObjectModel

Namespace ViewModel
  Public Class TreeViewItemViewModel
    Inherits ActionableViewModel

#Region " Data "

    ReadOnly _children As ObservableCollection(Of TreeViewItemViewModel)
    ReadOnly _parent As TreeViewItemViewModel

    Private _isExpanded As Boolean
    Private _isSelected As Boolean

#End Region

#Region " Dummy "
    Shared ReadOnly DummyChild As New TreeViewItemViewModel()
    'Private Shared _dummyCommand As RelayCommand
    'Public Shared ReadOnly Property DummyCommand As RelayCommand
    ' Get
    '  If _dummyCommand Is Nothing Then
    '   _dummyCommand = New RelayCommand(Sub(param) DummyCommandHandler(param))
    '  End If
    '  Return _dummyCommand
    ' End Get
    'End Property

    'Protected Shared Sub DummyCommandHandler(param As Object)
    'End Sub
#End Region

#Region " Constructors "

    Protected Sub New(parent As TreeViewItemViewModel, lazyLoadChildren As Boolean)
      MyBase.New("Item")
      _parent = parent

      _children = New ObservableCollection(Of TreeViewItemViewModel)()

      If lazyLoadChildren Then
        _children.Add(DummyChild)
      Else
        LoadChildren()
      End If

    End Sub

    ' This is used to create the DummyChild instance.
    Private Sub New()
      MyBase.New("Item")
    End Sub

#End Region

#Region " Presentation Members "

#Region " Children "

    ''' <summary>
    ''' Returns the logical child items of this object.
    ''' </summary>
    Public ReadOnly Property Children() As ObservableCollection(Of TreeViewItemViewModel)
      Get
        Return _children
      End Get
    End Property

#End Region

#Region " HasLoadedChildren "

    ''' <summary>
    ''' Returns true if this object's Children have not yet been populated.
    ''' </summary>
    Public ReadOnly Property HasDummyChild() As Boolean
      Get
        Return Me.Children.Count = 1 AndAlso Me.Children(0) Is DummyChild
      End Get
    End Property

#End Region

#Region " IsExpanded "

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

        ' Lazy load the child items, if necessary.
        If Me.HasDummyChild Then
          Me.Children.Remove(DummyChild)
          Me.LoadChildren()
        End If
      End Set
    End Property

#End Region

#Region " IsSelected "

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

#Region " Other "
    Public Overridable ReadOnly Property Key As String
      Get
        Return ""
      End Get
    End Property
    Public Overridable ReadOnly Property Name As String
      Get
        Return ""
      End Get
    End Property
    Public Overridable ReadOnly Property Image As ImageSource
      Get
        Return Nothing
      End Get
    End Property
#End Region

#Region " LoadChildren "

    ''' <summary>
    ''' Invoked when the child items need to be loaded on demand.
    ''' Subclasses can override this to populate the Children collection.
    ''' </summary>
    Protected Overridable Sub LoadChildren()
    End Sub

#End Region

#Region " Parent "

    Public ReadOnly Property Parent() As TreeViewItemViewModel
      Get
        Return _parent
      End Get
    End Property

#End Region

#End Region

  End Class
End Namespace
