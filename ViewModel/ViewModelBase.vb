Imports System.ComponentModel

Namespace ViewModel
 Public MustInherit Class ViewModelBase
  Implements INotifyPropertyChanged
  Implements IDisposable

#Region " Properties "
  Protected Property ThrowOnInvalidPropertyName As Boolean

  Private _DisplayName As String = ""
  Public Property DisplayName() As String
   Get
    Return _DisplayName
   End Get
   Set(ByVal value As String)
    _DisplayName = value
    Me.OnPropertyChanged("DisplayName")
   End Set
  End Property

  Public Property ID As String
#End Region

#Region " INotifyPropertyChanged "
  Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Protected Overridable Sub OnPropertyChanged(propertyName As String)
   Me.VerifyPropertyName(propertyName)
   RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
  End Sub

  <Conditional("DEBUG")>
  <DebuggerStepThrough()>
  Public Sub VerifyPropertyName(propertyName As String)
   ' Verify that the property name matches a real,  
   ' public, instance property on this object.
   If TypeDescriptor.GetProperties(Me)(propertyName) Is Nothing Then
    Dim msg As String = "Invalid property name: " & propertyName
    If Me.ThrowOnInvalidPropertyName Then
     Throw New Exception(msg)
    Else
     Debug.Fail(msg)
    End If
   End If
  End Sub
#End Region

#Region " IDisposable Support "
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
   If Not Me.disposedValue Then
    If disposing Then
     ' TODO: dispose managed state (managed objects).
    End If

    ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
    ' TODO: set large fields to null.
   End If
   Me.disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
   ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
   Dispose(True)
   GC.SuppressFinalize(Me)
  End Sub
#End Region

 End Class
End Namespace
