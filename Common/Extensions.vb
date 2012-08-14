Imports System.Windows.Interop

Namespace Common
 Public Module Extensions

  <System.Runtime.CompilerServices.Extension()> _
  Public Function GetIWin32Window(visual As System.Windows.Media.Visual) As System.Windows.Forms.IWin32Window
   Dim source = TryCast(System.Windows.PresentationSource.FromVisual(visual), System.Windows.Interop.HwndSource)
   Dim win As System.Windows.Forms.IWin32Window = New OldWindow(source.Handle)
   Return win
  End Function

  <System.Runtime.CompilerServices.Extension()> _
  Public Sub Refresh(uiElement As UIElement)
   uiElement.Dispatcher.Invoke(Windows.Threading.DispatcherPriority.Render, EmptyDelegate)
  End Sub
  Private EmptyDelegate As Action = Function() As action
                                     Return Nothing
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
