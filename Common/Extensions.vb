Imports System.Windows.Interop

Namespace Common
 Public Module Extensions

#Region " Various "
  <System.Runtime.CompilerServices.Extension()>
  Public Function GetIWin32Window(visual As System.Windows.Media.Visual) As System.Windows.Forms.IWin32Window
   Dim source = TryCast(System.Windows.PresentationSource.FromVisual(visual), System.Windows.Interop.HwndSource)
   Dim win As System.Windows.Forms.IWin32Window = New OldWindow(source.Handle)
   Return win
  End Function

  <System.Runtime.CompilerServices.Extension()>
  Public Sub Refresh(uiElement As UIElement)
   uiElement.Dispatcher.Invoke(Windows.Threading.DispatcherPriority.Render, EmptyDelegate)
  End Sub
  Private EmptyDelegate As Action = Function() As action
                                     Return Nothing
                                    End Function

  <System.Runtime.CompilerServices.Extension()>
  Public Function ToTree(flatList As List(Of String)) As TreeItem
   Dim rootItem As New TreeItem("Root", "")
   For Each item As String In flatList
    Dim addTo As TreeItem = rootItem
    For Each s As String In item.Split("\"c)
     If s <> "" Then
      If addTo.HasChild(s) Then
       addTo = addTo.GetChild(s)
      Else
       addTo = addTo.AddChild(s)
      End If
     End If
    Next
   Next
   Return rootItem
  End Function

  <System.Runtime.CompilerServices.Extension()>
  Public Function GetDataGridRows(grid As System.Windows.Controls.DataGrid) As IEnumerable(Of System.Windows.Controls.DataGridRow)
   Dim itemsSource = TryCast(grid.ItemsSource, IEnumerable)
   Dim res As New List(Of System.Windows.Controls.DataGridRow)
   If itemsSource Is Nothing Then Return res
   For Each item As Object In itemsSource
    Dim row = TryCast(grid.ItemContainerGenerator.ContainerFromItem(item), System.Windows.Controls.DataGridRow)
    If row IsNot Nothing Then
     res.Add(row)
    End If
   Next
   Return res
  End Function

  <System.Runtime.CompilerServices.Extension()>
  Public Function FindElementByName(root As FrameworkElement, name As String) As FrameworkElement
   Dim tree As New Stack(Of FrameworkElement)()
   tree.Push(root)
   While tree.Count > 0
    Dim current As FrameworkElement = tree.Pop()
    ' root is null
    If current.Name = name Then
     Return current
    End If
    Dim count As Integer = VisualTreeHelper.GetChildrenCount(current)
    For SupplierCounter As Integer = 0 To count - 1
     Dim child As DependencyObject = VisualTreeHelper.GetChild(current, SupplierCounter)
     If TypeOf child Is FrameworkElement Then
      tree.Push(DirectCast(child, FrameworkElement))
     End If
    Next
   End While
   Return Nothing
  End Function

  <System.Runtime.CompilerServices.Extension()>
  Public Function GetVisualChild(Of T As Visual)(parent As Visual, name As String) As T
   Dim child As T = Nothing
   Dim numVisuals As Integer = VisualTreeHelper.GetChildrenCount(parent)
   For i As Integer = 0 To numVisuals - 1
    Dim v As Visual = DirectCast(VisualTreeHelper.GetChild(parent, i), Visual)
    Dim element As FrameworkElement = TryCast(v, FrameworkElement)
    If element IsNot Nothing AndAlso element.Name.ToLower = name.ToLower Then
     child = TryCast(v, T)
    End If
    If child Is Nothing Then
     child = GetVisualChild(Of T)(v, name)
    End If
    If child IsNot Nothing Then
     Exit For
    End If
   Next
   Return child
  End Function
#End Region

#Region " Xml "
  <System.Runtime.CompilerServices.Extension()>
  Public Function CreateAndAppendElement(ByRef node As Xml.XmlNode, name As String) As Xml.XmlNode
   Dim newElement As Xml.XmlNode = node.OwnerDocument.CreateElement(name)
   node.AppendChild(newElement)
   Return newElement
  End Function

  <System.Runtime.CompilerServices.Extension()>
  Public Sub CreateAndAppendElement(ByRef node As Xml.XmlNode, name As String, value As String)
   Dim newElement As Xml.XmlNode = node.OwnerDocument.CreateElement(name)
   newElement.InnerText = value
   node.AppendChild(newElement)
  End Sub

  <System.Runtime.CompilerServices.Extension()>
  Public Sub AppendAttribute(ByRef node As Xml.XmlNode, name As String, value As String)
   Dim att As Xml.XmlAttribute = node.OwnerDocument.CreateAttribute(name)
   att.InnerText = value
   node.Attributes.Append(att)
  End Sub

  <System.Runtime.CompilerServices.Extension()>
  Public Function Replace(originalString As String, oldValue As String, newValue As String, comparisonType As StringComparison) As String
   If oldValue = "" Or oldValue = newValue Then Return originalString
   Dim startIndex As Integer = 0
   While True
    startIndex = originalString.IndexOf(oldValue, startIndex, comparisonType)
    If startIndex = -1 Then
     Exit While
    End If
    originalString = originalString.Substring(0, startIndex) & newValue & originalString.Substring(startIndex + oldValue.Length)
    startIndex += newValue.Length
   End While
   Return originalString
  End Function
#End Region

 End Module

#Region " OldWindow "
 Public Class OldWindow
  Implements System.Windows.Forms.IWin32Window
  Private ReadOnly _handle As System.IntPtr
  Public Sub New(handle As System.IntPtr)
   _handle = handle
  End Sub

  Private ReadOnly Property System_Windows_Forms_IWin32Window_Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
   Get
    Return _handle
   End Get
  End Property

 End Class
#End Region

End Namespace
