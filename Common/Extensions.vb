Imports System.Windows.Interop

Namespace Common
 Public Module Extensions

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

 End Module

 Public Class OldWindow
  Implements System.Windows.Forms.IWin32Window
  Private ReadOnly _handle As System.IntPtr
  Public Sub New(handle As System.IntPtr)
   _handle = handle
  End Sub

#Region "IWin32Window Members"
  Private ReadOnly Property System_Windows_Forms_IWin32Window_Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
   Get
    Return _handle
   End Get
  End Property
#End Region

 End Class
End Namespace
