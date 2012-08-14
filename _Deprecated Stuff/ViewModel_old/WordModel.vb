Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports RibbonSamples.Common
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Namespace ViewModel
 Public NotInheritable Class WordModel
  Private Sub New()
  End Sub
#Region "Application Menu"

  Public Shared ReadOnly Property [New]() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_New"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Open() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Open"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Open (Ctrl+O)"

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Open_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Save() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Save"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Save (Ctrl+S)"

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Save_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAs() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Save_As"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "SaveAs (F12)"

						Dim splitMenuItemData As New ApplicationSplitMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Save_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = splitMenuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAsWordDocument() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Word Document"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Save the file as a Word Document."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Document_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "W" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAsWordTemplate() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Word _Template"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Save the docuemt as a template that can be used to format future documents."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Document_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "T" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAsWord97To2003Document() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Word _97-2003 Document"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Save a copy of the document that is fully compatible with Word 97-2003."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Document_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "9" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAsOpenDocumentText() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Open_Document Text"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Save the document in the Open Document Format."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Document_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "D" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAsPdfOrXps() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_PDF or XPS"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Publish a copy of the document as a PDF or XPS file."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Document_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "P" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SaveAsOtherFormat() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Other Formats"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Open the Save As dialog box to select from all possible file types."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Save_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "O" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Print() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Print"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Print (Ctrl+P)"

						Dim splitMenuItemData As New ApplicationSplitMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Print_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = splitMenuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property PrintOptions() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Print Options"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Select a printer, number of copies, and other printing options before printing."

						Dim splitMenuItemData As New ApplicationSplitMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Print_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "P" _
						}
      _dataCollection(Str) = splitMenuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property QuickPrint() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Quick Print"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Send the document directly to the default printer without any changes."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Print_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "Q" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property PrintPreview() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Print Pre_view"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Preview and make changes to the pages before printing."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/PrintPreview_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "V" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Prepare() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Pr_epare"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Properties() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Properties"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "View and edit document properties, such as Title, Author and Keywords."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "P" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property InspectDocument() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Inspect Document"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Check the document for hidden metadata or personal information."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "I" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property EncryptDocument() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Encrypt Document"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Increase the security of the document by adding encryption."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "E" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RestrictPermission() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Restrict Permission"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Grant people access while restricting their ability to edit, copy and print."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewPermission_32x32.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "R" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property UnrestrictedAccess() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Unrestricted Access"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True, _
							Key .IsChecked = True, _
							Key .KeyTip = "U" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RestrictedAccess() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Restricted Access"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True, _
							Key .KeyTip = "R" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DoNotReplyAll() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Do Not Reply All"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MicrosoftAllAllRights() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Microsoft All - All Rights"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MicrosoftAllAllRightsExceptCopyAndPrint() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Microsoft All - All Rights Except Copy and Print"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MicrosoftAllReadOnly() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Microsoft All - Read Only"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MicrosoftFteAllRightsExceptCopyAndPrint() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Microsoft FTE - All Rights Except Copy and Print"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MicrosoftFteReadOnly() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Microsoft FTE - Read Only"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ManageCredentials() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Manage Credentials"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsCheckable = True, _
							Key .KeyTip = "M" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property AddADigitalSignature() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Add a Digital _Signature"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Ensure the integrity of the document by adding an invisible signature."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "S" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MarkAsFinal() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Mark as _Final"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Let the readers know the document is final and make it read-only."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "F" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RunCompatibilityChecker() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Run _Compatibility Checker"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Check for features not supported by earlier versions of Word."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "C" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Send() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Sen_d"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/SendLinkByEmail_32x32.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Email() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_E-mail"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Send a copy of the document in an email as an attachment."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/SendLinkByEmail_32x32.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "E" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property EmailAsPdfAttachment() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "E-mail as _PDF Attachment"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Send a copy of the document in a message as a PDF attachment."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/SendLinkByEmail_32x32.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "P" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property EmailAsXpsAttachment() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "E-mail as XP_S Attachment"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Send a copy of the document in a message as a XPS attachment."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/SendLinkByEmail_32x32.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "S" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property InternetFax() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Internet Fa_x"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Use an Internet fax service to fax the document."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Printer_48x48.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "X" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Publish() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "P_ublish"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/PublishPlan_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Blog() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Blog"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Create a new blog post with the content of the document."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ConnectionManager_48x48.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "B" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DocumentManagementServer() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Document Management Server"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Share the document by saving it to a document management server."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ConnectionManager_48x48.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "D" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property CreateDocumentWorkspace() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Create Document Workspace"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Create a new site for the document and keep the local copy synchronized."

						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ConnectionManager_48x48.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "C" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Workflows() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Wor_kflows"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NewDocument_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Close() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Close"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New ApplicationMenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FolderClose_48x48.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RecentDocuments() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Recent Documents"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim galleryCategoryData As New GalleryCategoryData(Of RecentDocumentData)()

      For i As Integer = 0 To 5
							Dim recentDocumentData As New RecentDocumentData() With { _
								Key .Index = i + 1, _
								Key .Label = "Recent Doc " & i _
							}
       galleryCategoryData.GalleryItemDataCollection.Add(recentDocumentData)
      Next

      _dataCollection(Str) = galleryCategoryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property WordOptions() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Word Options"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim buttonData As New ButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Options_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "I" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ExitWord() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Exit Word"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim buttonData As New ButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Close_16x16.png", UriKind.Relative), _
							Key .Command = ApplicationCommands.Close, _
							Key .KeyTip = "X" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#End Region


#Region "Help Button"

  Public Shared ReadOnly Property Help() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Help"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = "Help (F1)", _
							Key .ToolTipDescription = "Microsoft Ribbon for WPF" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#End Region

  Public Shared ReadOnly Property Clipboard() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Clipboard"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New GroupData(Str) With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative), _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative), _
							Key .KeyTip = "ZC" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#Region "Font Group Model"

  Public Shared ReadOnly Property Font() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New GroupData(Str) With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Font_16x16.png", UriKind.Relative), _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Font_32x32.png", UriKind.Relative), _
							Key .KeyTip = "ZF" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Paste() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Paste"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Paste (Ctrl+V)"
      Dim ToolTipDescription As String = "Paste the contents of the Clipboard."
      Dim DropDownToolTipDescription As String = "Click here for more options such as pasting only the values or formatting."

						Dim splitButtonData As New SplitButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative), _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Paste_32x32.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = ApplicationCommands.Paste, _
							Key .KeyTip = "V" _
						}
      splitButtonData.DropDownButtonData.ToolTipTitle = TooTipTitle
      splitButtonData.DropDownButtonData.ToolTipDescription = DropDownToolTipDescription
      splitButtonData.DropDownButtonData.Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute)
      splitButtonData.DropDownButtonData.KeyTip = "V"
      _dataCollection(Str) = splitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property PasteSpecial() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Paste _Special"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Paste Special (Alt+Ctrl+V)"

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property PasteAsHyperlink() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Paste As _Hyperlink"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Paste_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Cut() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Cut"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim CutToolTipTitle As String = "Cut (Ctrl+X)"
      Dim CutToolTipDescription As String = "Cut the selection and put it on the Clipboard."

						Dim buttonData As New ButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Cut_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = CutToolTipTitle, _
							Key .ToolTipDescription = CutToolTipDescription, _
							Key .Command = ApplicationCommands.Cut, _
							Key .KeyTip = "X" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Copy() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Copy"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Copy (Ctrl+C)"
      Dim ToolTipDescription As String = "Copy selection and put it on the Clipboard."

						Dim buttonData As New ButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Copy_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = ApplicationCommands.Copy, _
							Key .KeyTip = "C" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FormatPainter() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Format Painter"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Format Painter (Ctrl+Shift+C)"
      Dim ToolTipDescription As String = "Copy the formatting from one place and apply it to another. " & vbLf & vbLf & " Double click this button to apply the same formatting to multiple places in the document."

						Dim buttonData As New ButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FormatPainter_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = HelpFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "FP" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FontFace() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font Face"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Font (Ctrl+Shift+F)"
      Dim ToolTipDescription As String = "Change the font face."

						Dim comboBoxData As New ComboBoxData() With { _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .KeyTip = "FF" _
						}

      _dataCollection(Str) = comboBoxData
     End If
     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FontFaceGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font Face Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim galleryData As New GalleryData(Of FontFamily)() With { _
							Key .SelectedItem = SystemFonts.MessageFontFamily _
						}

						Dim recentlyUsedCategoryData As New GalleryCategoryData(Of FontFamily)() With { _
							Key .Label = "Recently Used Fonts" _
						}

      galleryData.CategoryDataCollection.Add(recentlyUsedCategoryData)

						Dim allFontsCategoryData As New GalleryCategoryData(Of FontFamily)() With { _
							Key .Label = "All Fonts" _
						}

      For Each fontFamily As FontFamily In System.Windows.Media.Fonts.SystemFontFamilies
       allFontsCategoryData.GalleryItemDataCollection.Add(fontFamily)
      Next

      galleryData.CategoryDataCollection.Add(allFontsCategoryData)

      Dim ChangeFontFace As Action(Of FontFamily) = Sub(parameter As FontFamily)
                                                     Dim wordControl__1 As UserControlWord = WordControl
                                                     If wordControl__1 IsNot Nothing Then
                                                      wordControl__1.ChangeFontFace(parameter)

                                                      If Not recentlyUsedCategoryData.GalleryItemDataCollection.Contains(parameter) Then
                                                       recentlyUsedCategoryData.GalleryItemDataCollection.Add(parameter)
                                                      End If
                                                     End If

                                                    End Sub

      Dim CanChangeFontFace As Func(Of FontFamily, Boolean) = Sub(parameter As FontFamily)
                                                               Dim wordControl__1 As UserControlWord = WordControl
                                                               If wordControl__1 IsNot Nothing Then
                                                                Return wordControl__1.CanChangeFontFace(parameter)
                                                               End If

                                                               Return False

                                                              End Sub

      Dim PreviewFontFace As Action(Of FontFamily) = Sub(parameter As FontFamily)
                                                      Dim wordControl__1 As UserControlWord = WordControl
                                                      If wordControl__1 IsNot Nothing Then
                                                       wordControl__1.PreviewFontFace(parameter)
                                                      End If

                                                     End Sub

      Dim CancelPreviewFontFace As Action = Sub()
                                             Dim wordControl__1 As UserControlWord = WordControl
                                             If wordControl__1 IsNot Nothing Then
                                              wordControl__1.CancelPreviewFontFace()
                                             End If

                                            End Sub

      galleryData.Command = New PreviewDelegateCommand(Of FontFamily)(ChangeFontFace, CanChangeFontFace, PreviewFontFace, CancelPreviewFontFace)

      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FontSize() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font Size"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Font Size (Ctrl+Shift+P)"
      Dim ToolTipDescription As String = "Change the font size."

						Dim comboBoxData As New ComboBoxData() With { _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .KeyTip = "FS" _
						}

      _dataCollection(Str) = comboBoxData
     End If
     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FontSizeGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font Size Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ChangeFontSize As Action(Of System.Nullable(Of Double)) = Sub(parameter As System.Nullable(Of Double))
                                                                     Dim wordControl__1 As UserControlWord = WordControl
                                                                     If wordControl__1 IsNot Nothing Then
                                                                      wordControl__1.ChangeFontSize(parameter)
                                                                     End If

                                                                    End Sub

      Dim CanChangeFontSize As Func(Of System.Nullable(Of Double), Boolean) = Sub(parameter As System.Nullable(Of Double))
                                                                               Dim wordControl__1 As UserControlWord = WordControl
                                                                               If wordControl__1 IsNot Nothing Then
                                                                                Return wordControl__1.CanChangeFontSize(parameter)
                                                                               End If

                                                                               Return False

                                                                              End Sub

      Dim PreviewFontSize As Action(Of System.Nullable(Of Double)) = Sub(parameter As System.Nullable(Of Double))
                                                                      Dim wordControl__1 As UserControlWord = WordControl
                                                                      If wordControl__1 IsNot Nothing Then
                                                                       wordControl__1.PreviewFontSize(parameter)
                                                                      End If

                                                                     End Sub

      Dim CancelPreviewFontSize As Action = Sub()
                                             Dim wordControl__1 As UserControlWord = WordControl
                                             If wordControl__1 IsNot Nothing Then
                                              wordControl__1.CancelPreviewFontSize()
                                             End If

                                            End Sub

						Dim galleryData As New GalleryData(Of System.Nullable(Of Double))() With { _
							Key .Command = New PreviewDelegateCommand(Of System.Nullable(Of Double))(ChangeFontSize, CanChangeFontSize, PreviewFontSize, CancelPreviewFontSize), _
							Key .SelectedItem = 11 _
						}

      Dim galleryCategoryData As New GalleryCategoryData(Of System.Nullable(Of Double))()
      galleryCategoryData.GalleryItemDataCollection.Add(8)
      galleryCategoryData.GalleryItemDataCollection.Add(9)
      galleryCategoryData.GalleryItemDataCollection.Add(10)
      galleryCategoryData.GalleryItemDataCollection.Add(11)
      galleryCategoryData.GalleryItemDataCollection.Add(12)
      galleryCategoryData.GalleryItemDataCollection.Add(14)
      galleryCategoryData.GalleryItemDataCollection.Add(16)
      galleryCategoryData.GalleryItemDataCollection.Add(18)
      galleryCategoryData.GalleryItemDataCollection.Add(20)
      galleryCategoryData.GalleryItemDataCollection.Add(22)
      galleryCategoryData.GalleryItemDataCollection.Add(24)
      galleryCategoryData.GalleryItemDataCollection.Add(28)
      galleryCategoryData.GalleryItemDataCollection.Add(36)
      galleryCategoryData.GalleryItemDataCollection.Add(48)
      galleryCategoryData.GalleryItemDataCollection.Add(72)

      galleryData.CategoryDataCollection.Add(galleryCategoryData)

      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property GrowFont() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Grow Font"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Grow Font (Ctrl+>)"
      Dim ToolTipDescription As String = "Increase the font size."

						Dim buttonData As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Font_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.IncreaseFontSize, _
							Key .KeyTip = "FG" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ShrinkFont() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Shrink Font"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Shrink Font (Ctrl+<)"
      Dim ToolTipDescription As String = "Decrease the font size."

						Dim buttonData As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Font_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.DecreaseFontSize, _
							Key .KeyTip = "FK" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ClearFormatting() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Clear Formatting"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Clear Formatting"
      Dim ToolTipDescription As String = "Clear all the formatting from the selection, leaving only the plain text."

						Dim buttonData As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ClearFormatting_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = HelpFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "E" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Bold() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Bold"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Bold (Ctrl+B)"
      Dim ToolTipDescription As String = "Make the selected text bold."

						Dim toggleButtonData As New ToggleButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Bold_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.ToggleBold, _
							Key .KeyTip = "1" _
						}
      _dataCollection(Str) = toggleButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Italic() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Italic"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Italic (Ctrl+I)"
      Dim ToolTipDescription As String = "Italicize the selected text."

						Dim toggleButtonData As New ToggleButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Italic_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.ToggleItalic, _
							Key .KeyTip = "2" _
						}
      _dataCollection(Str) = toggleButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Underline() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Underline"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Underline (Ctrl+U)"
      Dim ToolTipDescription As String = "Underline the selected text."

						Dim splitButtonData As New SplitButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/LineColor_16x16.png", UriKind.Relative), _
							Key .Command = EditingCommands.ToggleUnderline, _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .KeyTip = "3" _
						}

      _dataCollection(Str) = splitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property UnderlineGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Underline Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim galleryData As New GalleryData(Of String)() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/LineColor_16x16.png", UriKind.Relative), _
							Key .Command = EditingCommands.ToggleUnderline _
						}

      Dim galleryCategoryData As New GalleryCategoryData(Of String)()
      galleryCategoryData.GalleryItemDataCollection.Add("Underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Double underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Thick underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Dotted underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Dashed underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Dot-dash underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Dot-dot-dash underline")
      galleryCategoryData.GalleryItemDataCollection.Add("Wave underline")
      galleryData.CategoryDataCollection.Add(galleryCategoryData)

      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MoreUnderlines() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_More Underlines"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Color_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Strikethrough() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Strikethrough"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Strikethrough"
      Dim ToolTipDescription As String = "Draw a line through the middle of the selected text."

						Dim toggleButtonData As New ToggleButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Erase_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "4" _
						}
      _dataCollection(Str) = toggleButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Subscript() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Subscript"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Subscript (Ctrl+=)"
      Dim ToolTipDescription As String = "Create small letters below the test baseline."

						Dim toggleButtonData As New ToggleButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FontScript_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.ToggleSubscript, _
							Key .KeyTip = "5" _
						}
      _dataCollection(Str) = toggleButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Superscript() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Superscript"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Superscript (Ctrl+Shift++)"
      Dim ToolTipDescription As String = "Create small letters above the line of text. " & vbLf & vbLf & " To create a footnote, click Insert Footnote on References tab."

						Dim toggleButtonData As New ToggleButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FontScript_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.ToggleSuperscript, _
							Key .KeyTip = "6" _
						}
      _dataCollection(Str) = toggleButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ChangeCase() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Change Case"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Change Case"
      Dim ToolTipDescription As String = "Change all the selected text to UPPERCASE, lowercase or other common capitalizations."

						_dataCollection(Str) = New MenuButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FontScript_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = HelpFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "7" _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SentenceCase() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Sentence case."

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Uppercase() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_UPPERCASE"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Lowercase() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_lowercase"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property CapitalizeEachWord() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Capitalize Each Word"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ToggleCase() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_tOGGLE cASE"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property TextHighlightColor() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Text Highlight Color"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Text Highlight Color"
      Dim ToolTipDescription As String = "Make the text look like it was marked with a highlighter pen."

						Dim splitMenuItemData As New SplitMenuItemData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Highlight_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New PreviewDelegateCommand(Of Brush)(AddressOf ChangeTextHighlightColor, AddressOf CanChangeTextHighlightColor, AddressOf PreviewTextHighlightColor, AddressOf CancelPreviewTextHighlightColor), _
							Key .KeyTip = "I" _
						}

      _dataCollection(Str) = splitMenuItemData
     End If


     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property TextHighlightColorGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Text Highlight Color Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim galleryData As New GalleryData(Of Brush)() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Highlight_16x16.png", UriKind.Relative), _
							Key .Command = New PreviewDelegateCommand(Of Brush)(AddressOf ChangeTextHighlightColor, AddressOf CanChangeTextHighlightColor, AddressOf PreviewTextHighlightColor, AddressOf CancelPreviewTextHighlightColor), _
							Key .SelectedItem = SystemColors.ControlBrush _
						}

      Dim galleryCategoryData As New GalleryCategoryData(Of Brush)()
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Yellow)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Green)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Turquoise)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Pink)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Blue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Red)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkBlue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Teal)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Green)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Violet)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkRed)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkOrange)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkSeaGreen)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Aqua)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Gray)

      galleryData.CategoryDataCollection.Add(galleryCategoryData)

						galleryCategoryData = New GalleryCategoryData(Of Brush)() With { _
							Key .Label = "No Color" _
						}
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Transparent)
      galleryData.CategoryDataCollection.Add(galleryCategoryData)

      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Private Shared Sub ChangeTextHighlightColor(parameter As Brush)
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.ChangeTextHighlightColor(parameter)
   End If
  End Sub

  Private Shared Function CanChangeTextHighlightColor(parameter As Brush) As Boolean
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    Return wordControl__1.CanChangeTextHighlightColor(parameter)
   End If

   Return False
  End Function

  Private Shared Sub PreviewTextHighlightColor(parameter As Brush)
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.PreviewTextHighlightColor(parameter)
   End If
  End Sub

  Private Shared Sub CancelPreviewTextHighlightColor()
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.CancelPreviewTextHighlightColor()
   End If
  End Sub

  Public Shared ReadOnly Property StopHighlighting() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Stop Highlighting"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FontColor() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font Color"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Font Color"
      Dim ToolTipDescription As String = "Change the text color."

						Dim splitMenuItemData As New SplitMenuItemData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FontColor_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = HelpFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New PreviewDelegateCommand(Of Brush)(AddressOf ChangeFontColor, AddressOf CanChangeFontColor, AddressOf PreviewFontColor, AddressOf CancelPreviewFontColor), _
							Key .KeyTip = "FC" _
						}

      _dataCollection(Str) = splitMenuItemData
     End If


     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FontColorGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Font Color Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim galleryData As New GalleryData(Of Brush)() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/FontColor_16x16.png", UriKind.Relative), _
							Key .Command = New PreviewDelegateCommand(Of Brush)(AddressOf ChangeFontColor, AddressOf CanChangeFontColor, AddressOf PreviewFontColor, AddressOf CancelPreviewFontColor), _
							Key .SelectedItem = SystemColors.ControlBrush _
						}

						Dim galleryCategoryData As New GalleryCategoryData(Of Brush)() With { _
							Key .Label = "Automatic Color" _
						}
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Black)
      galleryData.CategoryDataCollection.Add(galleryCategoryData)

						galleryCategoryData = New GalleryCategoryData(Of Brush)() With { _
							Key .Label = "Theme Colors" _
						}
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.White)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Black)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Tan)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkBlue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Blue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Red)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Olive)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Purple)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Aqua)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Orange)

      Dim percent10 As Single = 0.9F
      Dim percent25 As Single = 0.75F
      Dim percent40 As Single = 0.6F
      Dim percent55 As Single = 0.45F
      Dim percent70 As Single = 0.3F

      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.White.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Black.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Tan.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.DarkBlue.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Blue.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Red.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Olive.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Purple.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Aqua.Color * percent10))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Orange.Color * percent10))

      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.White.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Black.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Tan.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.DarkBlue.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Blue.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Red.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Olive.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Purple.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Aqua.Color * percent25))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Orange.Color * percent25))

      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.White.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Black.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Tan.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.DarkBlue.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Blue.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Red.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Olive.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Purple.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Aqua.Color * percent40))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Orange.Color * percent40))

      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.White.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Black.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Tan.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.DarkBlue.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Blue.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Red.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Olive.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Purple.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Aqua.Color * percent55))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Orange.Color * percent55))

      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.White.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Black.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Tan.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.DarkBlue.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Blue.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Red.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Olive.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Purple.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Aqua.Color * percent70))
      galleryCategoryData.GalleryItemDataCollection.Add(New SolidColorBrush(Brushes.Orange.Color * percent70))

      galleryData.CategoryDataCollection.Add(galleryCategoryData)

						galleryCategoryData = New GalleryCategoryData(Of Brush)() With { _
							Key .Label = "Standard Colors" _
						}
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkRed)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Red)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Orange)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Yellow)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.LightGreen)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Green)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.LightBlue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Blue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.DarkBlue)
      galleryCategoryData.GalleryItemDataCollection.Add(Brushes.Purple)

      galleryData.CategoryDataCollection.Add(galleryCategoryData)

      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Private Shared Sub ChangeFontColor(parameter As Brush)
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.ChangeFontColor(parameter)
   End If
  End Sub

  Private Shared Function CanChangeFontColor(parameter As Brush) As Boolean
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    Return wordControl__1.CanChangeFontColor(parameter)
   End If

   Return False
  End Function

  Private Shared Sub PreviewFontColor(parameter As Brush)
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.PreviewFontColor(parameter)
   End If
  End Sub

  Private Shared Sub CancelPreviewFontColor()
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.CancelPreviewFontColor()
   End If
  End Sub

  Public Shared ReadOnly Property MoreColors() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_More Colors"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Color_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#End Region

#Region "Paragraph Group Model"

  Public Shared ReadOnly Property Paragraph() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Paragraph"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New GroupData(Str) With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Paragraph_16x16.png", UriKind.Relative), _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Paragraph_32x32.png", UriKind.Relative), _
							Key .KeyTip = "ZP" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Bullets() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Bullets"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Bullets"
      Dim ToolTipDescription As String = "Start a bulleted list." & vbLf & vbLf & "Click the arrow to choose different bullet styles."

						Dim splitButtonData As New SplitButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Bullets_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .IsCheckable = True, _
							Key .Command = EditingCommands.ToggleBullets, _
							Key .IsVerticallyResizable = True, _
							Key .IsHorizontallyResizable = True, _
							Key .KeyTip = "U" _
						}
      _dataCollection(Str) = splitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property BulletsGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Bullets Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim galleryData As New GalleryData(Of Uri)() From { _
      }

						Dim recentlyUsedCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Recently Used Bullets" _
						}

      galleryData.CategoryDataCollection.Add(recentlyUsedCategoryData)

						Dim bulletLibraryCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Bullet Library" _
						}
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/DownArrow_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/LeftArrow_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Minus_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Plus_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/RefreshArrow_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/RightArrow_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Tick_32x32.png", UriKind.Relative))
      bulletLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/UpArrow_32x32.png", UriKind.Relative))

      galleryData.CategoryDataCollection.Add(bulletLibraryCategoryData)

						Dim documentBulletsCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Document Bullets" _
						}

      documentBulletsCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Tick_32x32.png", UriKind.Relative))

      galleryData.CategoryDataCollection.Add(documentBulletsCategoryData)

						Dim galleryCommandExecuted As Action(Of Uri) = Sub(parameter As Uri) If Not recentlyUsedCategoryData.GalleryItemDataCollection.Contains(parameter) Then
      recentlyUsedCategoryData.GalleryItemDataCollection.Add(parameter)
     End If

     Dim galleryCommandCanExecute As Func(Of Uri, Boolean) = Function(parameter As Uri) True

     galleryData.Command = New DelegateCommand(Of Uri)(galleryCommandExecuted, galleryCommandCanExecute)
     _dataCollection(Str) = galleryData
					End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ChangeListLevel() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Change List Level"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/MultiLevelList_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ChangeListLevelGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "ChangeListLevelGallery"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim galleryData As New GalleryData(Of String)() From { _
      }

      Dim categoryData As New GalleryCategoryData(Of String)() From { _
      }

      categoryData.GalleryItemDataCollection.Add("  > First")
      categoryData.GalleryItemDataCollection.Add("    > Second")
      categoryData.GalleryItemDataCollection.Add("      > Third")
      categoryData.GalleryItemDataCollection.Add("        > Fourth")
      categoryData.GalleryItemDataCollection.Add("          > Fifth")
      categoryData.GalleryItemDataCollection.Add("            > Sixth")
      categoryData.GalleryItemDataCollection.Add("              > Seventh")
      categoryData.GalleryItemDataCollection.Add("                > Eighth")
      categoryData.GalleryItemDataCollection.Add("                  > Ninth")

      galleryData.CategoryDataCollection.Add(categoryData)

      galleryData.Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute)
      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DefaultNewBullet() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Default New Bullet..."

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DefaultNewNumberFormat() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Default New Number Format..."

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DefaultNewMultilevelList() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Default New Multilevel List..."

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DefaultNewListStyle() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Default New _List Style..."

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Default_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Numbering() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Numbering"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Numbering"
      Dim ToolTipDescription As String = "Start a numbered list." & vbLf & vbLf & "Click the arrow to choose different numbering formats."

						Dim splitButtonData As New SplitButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Numbering_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .IsCheckable = True, _
							Key .Command = EditingCommands.ToggleNumbering, _
							Key .IsVerticallyResizable = True, _
							Key .IsHorizontallyResizable = True, _
							Key .KeyTip = "N" _
						}
      _dataCollection(Str) = splitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property NumberingGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Numbering Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim galleryData As New GalleryData(Of Uri)() From { _
      }

						Dim recentlyUsedCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Recently Used Number Formats" _
						}

      galleryData.CategoryDataCollection.Add(recentlyUsedCategoryData)

						Dim numberingLibraryCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Numbering Library" _
						}
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/DownArrow_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/LeftArrow_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Minus_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Plus_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/RefreshArrow_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/RightArrow_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Tick_32x32.png", UriKind.Relative))
      numberingLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/UpArrow_32x32.png", UriKind.Relative))

      galleryData.CategoryDataCollection.Add(numberingLibraryCategoryData)

						Dim documentNumberFormatsCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Document Number Formats" _
						}

      documentNumberFormatsCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Tick_32x32.png", UriKind.Relative))

      galleryData.CategoryDataCollection.Add(documentNumberFormatsCategoryData)

						Dim galleryCommandExecuted As Action(Of Uri) = Sub(parameter As Uri) If Not recentlyUsedCategoryData.GalleryItemDataCollection.Contains(parameter) Then
      recentlyUsedCategoryData.GalleryItemDataCollection.Add(parameter)
     End If

     Dim galleryCommandCanExecute As Func(Of Uri, Boolean) = Function(parameter As Uri) True

     galleryData.Command = New DelegateCommand(Of Uri)(galleryCommandExecuted, galleryCommandCanExecute)
     _dataCollection(Str) = galleryData
					End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MultilevelList() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "MultilevelList"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Multilevel List"
      Dim ToolTipDescription As String = "Start a multilevel list." & vbLf & vbLf & "Click the arrow to choose different multilevel list styles."

						Dim menuButtonData As New MenuButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/MultiLevelList_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsVerticallyResizable = True, _
							Key .IsHorizontallyResizable = True, _
							Key .KeyTip = "M" _
						}
      _dataCollection(Str) = menuButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property MultilevelListGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "MultilevelList Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim galleryData As New GalleryData(Of Uri)() With { _
							Key .CanUserFilter = True _
						}

						Dim currentListCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Current List" _
						}

      galleryData.CategoryDataCollection.Add(currentListCategoryData)

						Dim listLibraryCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "List Library" _
						}
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/DownArrow_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/LeftArrow_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Minus_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Plus_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/RefreshArrow_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/RightArrow_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Tick_32x32.png", UriKind.Relative))
      listLibraryCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/UpArrow_32x32.png", UriKind.Relative))

      galleryData.CategoryDataCollection.Add(listLibraryCategoryData)

						Dim documentListsCategoryData As New GalleryCategoryData(Of Uri)() With { _
							Key .Label = "Lists in Current Documents" _
						}

      documentListsCategoryData.GalleryItemDataCollection.Add(New Uri("/RibbonWindowSample;component/Images/Tick_32x32.png", UriKind.Relative))

      galleryData.CategoryDataCollection.Add(documentListsCategoryData)

						Dim galleryCommandExecuted As Action(Of Uri) = Sub(parameter As Uri) If Not currentListCategoryData.GalleryItemDataCollection.Contains(parameter) Then
      currentListCategoryData.GalleryItemDataCollection.Clear()
      currentListCategoryData.GalleryItemDataCollection.Add(parameter)
     End If

     Dim galleryCommandCanExecute As Func(Of Uri, Boolean) = Function(parameter As Uri) True

     galleryData.Command = New DelegateCommand(Of Uri)(galleryCommandExecuted, galleryCommandCanExecute)
     _dataCollection(Str) = galleryData
					End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DecreaseIndent() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "DecreaseIndent"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Decrease Indent"
      Dim ToolTipDescription As String = "Decreases the indent level of the paragraph."

						Dim buttonData As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/DecreaseIndent_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.DecreaseIndentation, _
							Key .KeyTip = "AO" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property IncreaseIndent() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "IncreaseIndent"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Increase Indent"
      Dim ToolTipDescription As String = "Increases the indent level of the paragraph."

						Dim buttonData As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/IncreaseIndent_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.IncreaseIndentation, _
							Key .KeyTip = "AI" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Sort() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Sort"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Sort"
      Dim ToolTipDescription As String = "Alphabetize the selected text or sort numerical data."
      Dim ToolTipFooterTitle As String = "Press F1 for more help."

						Dim buttonData As New ButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Sort_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = ToolTipFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "SO" _
						}
      _dataCollection(Str) = buttonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ShowHide() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "ShowHide"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Show/Hide (Ctrl + *)"
      Dim ToolTipDescription As String = "Show paragraph marks and other hidden formatting symbols."
      Dim ToolTipFooterTitle As String = "Press F1 for more help."

						Dim toggleButtonData As New ToggleButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ShowHide_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = ToolTipFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "8" _
						}
      _dataCollection(Str) = toggleButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property AlignTextLeft() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "AlignTextLeft"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Align Text Left (Ctrl + L)"
      Dim ToolTipDescription As String = "Align text to the left."

						Dim radioButtonData As New RadioButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/AlignLeft_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.AlignLeft, _
							Key .KeyTip = "AL" _
						}
      _dataCollection(Str) = radioButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property AlignTextCenter() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "AlignTextCenter"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Center (Ctrl + E)"
      Dim ToolTipDescription As String = "Center text."

						Dim radioButtonData As New RadioButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/AlignCenter_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.AlignCenter, _
							Key .KeyTip = "AC" _
						}
      _dataCollection(Str) = radioButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property AlignTextRight() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "AlignTextRight"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Align Text Left (Ctrl + L)"
      Dim ToolTipDescription As String = "Align text to the right."

						Dim radioButtonData As New RadioButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/AlignRight_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.AlignRight, _
							Key .KeyTip = "AR" _
						}
      _dataCollection(Str) = radioButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Justify() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Justify"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Justify (Ctrl + J)"
      Dim ToolTipDescription As String = "Aligns text to both left and right margins, adding extra space between words as necessary." & vbLf & vbLf
      ToolTipDescription += "This creates a clean along left and right side of the page."

						Dim radioButtonData As New RadioButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Justify_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = EditingCommands.AlignJustify, _
							Key .KeyTip = "AJ" _
						}
      _dataCollection(Str) = radioButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacing() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacing"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Line Spacing"
      Dim ToolTipDescription As String = "Change the spacing between line of text." & vbLf & vbLf
      ToolTipDescription += "You can also customize the amount of space added before and after the paragraphs."

						Dim menuButtonData As New MenuButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/LineSpacing_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "K" _
						}
      _dataCollection(Str) = menuButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#Region "LineSpacing Model"

  Private Shared Sub SetIsCheckedOfLineSpacingMenuItem(menuItemData As MenuItemData, parameter As Object)
   menuItemData.IsChecked = (menuItemData Is parameter)
  End Sub

  Private Shared Sub LineSpacingMenuItemDefaultExecute(menuItemData As MenuItemData)
   SetIsCheckedOfLineSpacingMenuItem(DirectCast(LineSpacingFirstValue, MenuItemData), menuItemData)
   SetIsCheckedOfLineSpacingMenuItem(DirectCast(LineSpacingSecondValue, MenuItemData), menuItemData)
   SetIsCheckedOfLineSpacingMenuItem(DirectCast(LineSpacingThirdValue, MenuItemData), menuItemData)
   SetIsCheckedOfLineSpacingMenuItem(DirectCast(LineSpacingFourthValue, MenuItemData), menuItemData)
   SetIsCheckedOfLineSpacingMenuItem(DirectCast(LineSpacingFifthValue, MenuItemData), menuItemData)
   SetIsCheckedOfLineSpacingMenuItem(DirectCast(LineSpacingSixthValue, MenuItemData), menuItemData)
  End Sub

  Private Shared Function LineSpacingMenuItemDefaultCanExecute(menuItemData As MenuItemData) As Boolean
   Return True
  End Function

  Public Shared ReadOnly Property LineSpacingFirstValue() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacingFirstValue"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = "1.0", _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf LineSpacingMenuItemDefaultExecute, AddressOf LineSpacingMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacingSecondValue() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacingSecondValue"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = "1.15", _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf LineSpacingMenuItemDefaultExecute, AddressOf LineSpacingMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacingThirdValue() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacingThirdValue"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = "1.5", _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf LineSpacingMenuItemDefaultExecute, AddressOf LineSpacingMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacingFourthValue() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacingFourthValue"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = "2.0", _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf LineSpacingMenuItemDefaultExecute, AddressOf LineSpacingMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacingFifthValue() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacingFifthValue"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = "2.5", _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf LineSpacingMenuItemDefaultExecute, AddressOf LineSpacingMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacingSixthValue() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "LineSpacingSixthValue"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = "3.0", _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf LineSpacingMenuItemDefaultExecute, AddressOf LineSpacingMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LineSpacingOptions() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Line Spacing Options..."

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property AddSpaceBeforeParagraph() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Add Space _Before Paragraph"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/UpArrow_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RemoveSpaceAfterParagraph() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Remove Space _After Paragraph"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/DownArrow_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#End Region

  Public Shared ReadOnly Property Shading() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Shading"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Shading"
      Dim ToolTipDescription As String = "Color the background behind selected text or paragraph"
      Dim ToolTipFooterTitle As String = "Press F1 for more help."

						Dim splitButtonData As New SplitButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Shading_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = ToolTipFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "H" _
						}
      _dataCollection(Str) = splitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Borders() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Borders"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Borders"

						Dim splitButtonData As New SplitButtonData() With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Borders_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "B" _
						}
      _dataCollection(Str) = splitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#Region "Borders Model"

  Private Shared Sub BorderMenuItemDefaultExecute(menuItemData As MenuItemData)
   Dim bottomBorder__1 As MenuItemData = DirectCast(BottomBorder, MenuItemData)
   Dim topBorder__2 As MenuItemData = DirectCast(TopBorder, MenuItemData)
   Dim leftBorder__3 As MenuItemData = DirectCast(LeftBorder, MenuItemData)
   Dim rightBorder__4 As MenuItemData = DirectCast(RightBorder, MenuItemData)
   Dim noBorder__5 As MenuItemData = DirectCast(NoBorder, MenuItemData)
   Dim allBorders__6 As MenuItemData = DirectCast(AllBorders, MenuItemData)
   Dim outsideBorders__7 As MenuItemData = DirectCast(OutsideBorders, MenuItemData)
   Dim insideBorders__8 As MenuItemData = DirectCast(InsideBorders, MenuItemData)
   Dim insideHorizontalBorder__9 As MenuItemData = DirectCast(InsideHorizontalBorder, MenuItemData)
   Dim insideVerticalBorder__10 As MenuItemData = DirectCast(InsideVerticalBorder, MenuItemData)

   If menuItemData Is bottomBorder__1 OrElse menuItemData Is topBorder__2 OrElse menuItemData Is leftBorder__3 OrElse menuItemData Is rightBorder__4 Then
    outsideBorders__7.IsChecked = (bottomBorder__1.IsChecked AndAlso topBorder__2.IsChecked AndAlso leftBorder__3.IsChecked AndAlso rightBorder__4.IsChecked)
   End If

   If menuItemData Is outsideBorders__7 Then
    bottomBorder__1.IsChecked = InlineAssignHelper(topBorder__2.IsChecked, InlineAssignHelper(leftBorder__3.IsChecked, InlineAssignHelper(rightBorder__4.IsChecked, outsideBorders__7.IsChecked)))
   End If

   If menuItemData Is insideHorizontalBorder__9 OrElse menuItemData Is insideVerticalBorder__10 Then
    insideBorders__8.IsChecked = (insideHorizontalBorder__9.IsChecked AndAlso insideVerticalBorder__10.IsChecked)
   End If

   If menuItemData Is insideBorders__8 Then
    insideHorizontalBorder__9.IsChecked = InlineAssignHelper(insideVerticalBorder__10.IsChecked, insideBorders__8.IsChecked)
   End If

   If menuItemData Is noBorder__5 Then
    bottomBorder__1.IsChecked = False
    topBorder__2.IsChecked = False
    leftBorder__3.IsChecked = False
    rightBorder__4.IsChecked = False
    outsideBorders__7.IsChecked = False
    insideBorders__8.IsChecked = False
    insideHorizontalBorder__9.IsChecked = False
    insideVerticalBorder__10.IsChecked = False
   End If

   If menuItemData Is allBorders__6 Then
    bottomBorder__1.IsChecked = True
    topBorder__2.IsChecked = True
    leftBorder__3.IsChecked = True
    rightBorder__4.IsChecked = True
    outsideBorders__7.IsChecked = True
    insideBorders__8.IsChecked = True
    insideHorizontalBorder__9.IsChecked = True
    insideVerticalBorder__10.IsChecked = True
   End If
  End Sub

  Private Shared Function BordersMenuItemDefaultCanExecute(menuItemData As MenuItemData) As Boolean
   Return True
  End Function

  Public Shared ReadOnly Property BottomBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Bottom Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/BottomBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property TopBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "To_p Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/TopBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LeftBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Left Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/LeftBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property RightBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Right Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/RightBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property NoBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_No Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/NoBorder_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property AllBorders() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_All Borders"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Borders_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property InsideBorders() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Inside Borders"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/InsideBorders_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property OutsideBorders() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Out_side Borders"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/OuterBorders_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property InsideHorizontalBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Inside _Horizontal Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/InsideHorizontalBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property InsideVerticalBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Inside _Vertical Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/InsideVerticalBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DiagonalDownBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Diagonal Do_wn Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/DiagonalDownBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DiagonalUpBorder() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Diagonal _Up Border"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/DiagonalUpBorder_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(Of MenuItemData)(AddressOf BorderMenuItemDefaultExecute, AddressOf BordersMenuItemDefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property HorizontalLine() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Hori_zontal Line"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/InsideHorizontalBorder_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property DrawTable() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Draw Table"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/DrawTable_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ViewGridLines() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "View _Gridlines"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ShowGridlines_16x16.png", UriKind.Relative), _
							Key .IsCheckable = True, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property BordersAndShading() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "B_orders And Shading"

     If Not _dataCollection.ContainsKey(Str) Then
						_dataCollection(Str) = New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/BordersAndShading_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#End Region

#End Region

#Region "Styles Group Model"

  Public Shared ReadOnly Property Styles() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Styles"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New GroupData(Str) With { _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/StylesGroup.png", UriKind.Relative), _
							Key .KeyTip = "ZS" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property


  Public Shared ReadOnly Property ChangeStyles() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Change Styles"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Change Styles"
      Dim ToolTipDescription As String = "Change the set of styles, colors, fonts, and paragraph spacing used in this document."

						Dim menuButtonData As New MenuButtonData() With { _
							Key .Label = Str, _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Styles_32x32.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .KeyTip = "G" _
						}
      _dataCollection(Str) = menuButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property StylesSet() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Style Set"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Forecolor_16x16.png", UriKind.Relative), _
							Key .IsVerticallyResizable = True, _
							Key .KeyTip = "Y" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Colors() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Colors"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ChooseColor_16x16.png", UriKind.Relative), _
							Key .IsVerticallyResizable = True, _
							Key .KeyTip = "C" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Fonts() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Fonts"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Font_16x16.png", UriKind.Relative), _
							Key .IsVerticallyResizable = True, _
							Key .KeyTip = "F" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property ParagraphSpacing() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Paragraph Spacing"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/ParagraphSpacing_16x16.png", UriKind.Relative), _
							Key .KeyTip = "P" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SetAsDefault() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Set as Default"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = Str
      Dim ToolTipDescription As String = "Set the cuurent style set and theme as the default when you create a new document."

						Dim menuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "S" _
						}
      _dataCollection(Str) = menuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property StylesSetGalleryData() As GalleryData(Of StylesSet)
   Get
    SyncLock _lockObject
     Dim Str As String = "StylesSetGalleryData"

     If Not _dataCollection.ContainsKey(Str) Then
      ' TODO: replace string with an object (IsChecked, StyleName). Define DataTemplate
      Dim stylesData As New GalleryData(Of StylesSet)()
      Dim singleCategory As New GalleryCategoryData(Of StylesSet)()

						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Default (Black and White)", _
							Key .IsSelected = True _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Distinctive" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Elegant" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Fancy" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Formal" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Manuscript" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Modern" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Newsprint" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Perspective" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Simple" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Thatch" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Traditional" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Word 2003" _
						})
						singleCategory.GalleryItemDataCollection.Add(New StylesSet() With { _
							Key .Label = "Word 2010" _
						})

      stylesData.CategoryDataCollection.Clear()
      stylesData.CategoryDataCollection.Add(singleCategory)
      _dataCollection(Str) = stylesData
     End If

     Return TryCast(_dataCollection(Str), GalleryData(Of StylesSet))
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property StylesColorsGalleryData() As GalleryData(Of String)
   Get
    SyncLock _lockObject
     Dim Str As String = "StylesColorsGalleryData"

     If Not _dataCollection.ContainsKey(Str) Then
      ' TODO: replace string with an object (Color Palette, StyleName). Define DataTemplate
      Dim stylesData As New GalleryData(Of String)()
      Dim singleCategory As New GalleryCategoryData(Of String)()

      singleCategory.Label = "Built-In"
      singleCategory.GalleryItemDataCollection.Add("Office")
      singleCategory.GalleryItemDataCollection.Add("Grayscale")
      singleCategory.GalleryItemDataCollection.Add("Adjacency")
      singleCategory.GalleryItemDataCollection.Add("Angles")
      singleCategory.GalleryItemDataCollection.Add("Apex")
      singleCategory.GalleryItemDataCollection.Add("Apothecary")
      singleCategory.GalleryItemDataCollection.Add("Aspect")
      singleCategory.GalleryItemDataCollection.Add("Austin")
      singleCategory.GalleryItemDataCollection.Add("Black Tie")
      singleCategory.GalleryItemDataCollection.Add("Civic")
      singleCategory.GalleryItemDataCollection.Add("Clarity")
      singleCategory.GalleryItemDataCollection.Add("Composite")
      singleCategory.GalleryItemDataCollection.Add("Concourse")
      singleCategory.GalleryItemDataCollection.Add("Couture")
      singleCategory.GalleryItemDataCollection.Add("Elemental")
      singleCategory.GalleryItemDataCollection.Add("Equity")
      singleCategory.GalleryItemDataCollection.Add("Essential")
      singleCategory.GalleryItemDataCollection.Add("Executive")
      singleCategory.GalleryItemDataCollection.Add("Flow")
      singleCategory.GalleryItemDataCollection.Add("Foundry")
      singleCategory.GalleryItemDataCollection.Add("Grid")
      singleCategory.GalleryItemDataCollection.Add("Horizon")
      singleCategory.GalleryItemDataCollection.Add("Median")
      singleCategory.GalleryItemDataCollection.Add("Newsprint")
      singleCategory.GalleryItemDataCollection.Add("Perspective")
      singleCategory.GalleryItemDataCollection.Add("Solstice")
      singleCategory.GalleryItemDataCollection.Add("Technic")
      singleCategory.GalleryItemDataCollection.Add("Urban")
      singleCategory.GalleryItemDataCollection.Add("Verve")
      singleCategory.GalleryItemDataCollection.Add("Waveform")

      stylesData.CategoryDataCollection.Add(singleCategory)
      _dataCollection(Str) = stylesData
     End If

     Return TryCast(_dataCollection(Str), GalleryData(Of String))
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property StylesFontsGalleryData() As GalleryData(Of ThemeFonts)
   Get
    SyncLock _lockObject
     Dim Str As String = "StylesFontsGalleryData"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim stylesData As New GalleryData(Of ThemeFonts)()
      Dim singleCategory As New GalleryCategoryData(Of ThemeFonts)()

      singleCategory.Label = "Built-In"
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office 2", _
							Key .Field2 = "Calibri", _
							Key .Field3 = "Cambria", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office Classic", _
							Key .Field2 = "Arial", _
							Key .Field3 = "Times New Roman", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office Classic 2", _
							Key .Field2 = "Arial", _
							Key .Field3 = "Arial", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Adjacency", _
							Key .Field2 = "Franklin Gothic", _
							Key .Field3 = "Franklin Gothic Book", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Angles", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Apex", _
							Key .Field2 = "Lucida Sans", _
							Key .Field3 = "Book Antiqua", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Apothecary", _
							Key .Field2 = "Book Antiqua", _
							Key .Field3 = "Century Gothic", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Aspect", _
							Key .Field2 = "Verdana", _
							Key .Field3 = "Verdana", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Austin", _
							Key .Field2 = "Century Gothic", _
							Key .Field3 = "Century Gothic", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Black Tie", _
							Key .Field2 = "Garamond", _
							Key .Field3 = "Garamond", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Civic", _
							Key .Field2 = "Georgia", _
							Key .Field3 = "Georgia", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})
						singleCategory.GalleryItemDataCollection.Add(New ThemeFonts() With { _
							Key .Label = "Office", _
							Key .Field2 = "Cambria", _
							Key .Field3 = "Calibri", _
							Key .Field1 = "/RibbonWindowSample;component/Images/ThemeFonts.png" _
						})

      stylesData.CategoryDataCollection.Add(singleCategory)
      _dataCollection(Str) = stylesData
     End If

     Return TryCast(_dataCollection(Str), GalleryData(Of ThemeFonts))
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property StylesParagraphGalleryData() As GalleryData(Of String)
   Get
    SyncLock _lockObject
     Dim Str As String = "StylesParagraphGalleryData"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim stylesData As New GalleryData(Of String)()
      Dim firstCategory As New GalleryCategoryData(Of String)()
      firstCategory.Label = "Style Set"
      firstCategory.GalleryItemDataCollection.Add("Traditional")

      Dim secondCategory As New GalleryCategoryData(Of String)()
      secondCategory.Label = "Built-In"
      secondCategory.GalleryItemDataCollection.Add("No Paragraph Space")
      secondCategory.GalleryItemDataCollection.Add("Compact")
      secondCategory.GalleryItemDataCollection.Add("Tight")
      secondCategory.GalleryItemDataCollection.Add("Open")
      secondCategory.GalleryItemDataCollection.Add("Relaxed")
      secondCategory.GalleryItemDataCollection.Add("Double")

      stylesData.CategoryDataCollection.Add(firstCategory)
      stylesData.CategoryDataCollection.Add(secondCategory)
      _dataCollection(Str) = stylesData
     End If

     Return TryCast(_dataCollection(Str), GalleryData(Of String))
    End SyncLock
   End Get
  End Property

#End Region

#Region "Editing model"

  Public Shared ReadOnly Property Editing() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Editing"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New GroupData(Str) With { _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Find_16x16.png", UriKind.Relative), _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Find_32x32.png", UriKind.Relative), _
							Key .KeyTip = "ZN" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property Find() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Find"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Find (Ctrl+F)"
      Dim ToolTipDescription As String = "Find text or other content in the document."
      Dim DropDownToolTipDescription As String = "Find and select specific text, formatting or other type of information within the document."
      Dim DropDownToolTipFooter As String = "You can also replace the information with new text or formatting."

						Dim FindSplitButtonData As New SplitButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Find_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = TooTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "FD" _
						}
      FindSplitButtonData.DropDownButtonData.ToolTipTitle = TooTipTitle
      FindSplitButtonData.DropDownButtonData.ToolTipDescription = DropDownToolTipDescription
      FindSplitButtonData.DropDownButtonData.ToolTipFooterDescription = DropDownToolTipFooter
      FindSplitButtonData.DropDownButtonData.Command = New DelegateCommand(Sub()

                                                                           End Sub)
      _dataCollection(Str) = FindSplitButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FindMenuItem() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Find"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Find (Ctrl+F)"
      Dim ToolTipDescription As String = "Find text or other content in the document."

						Dim FindMenuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = TooTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Find_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "F" _
						}
      _dataCollection(Str) = FindMenuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property


  Public Shared ReadOnly Property AdvancedFind() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Advanced Find..."

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Advanced Find"
      Dim ToolTipDescription As String = "Find text in the document."

						Dim FindMenuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = TooTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Find_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "A" _
						}
      _dataCollection(Str) = FindMenuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property


  Public Shared ReadOnly Property [GoTo]() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Go To..."

     If Not _dataCollection.ContainsKey(Str) Then
      Dim TooTipTitle As String = "Go To (Ctrl+G)"
      Dim ToolTipDescription As String = "Navigate to a specific place in the document." & vbLf & vbCr & vbLf & vbCr & "Depending on the type of the document, you can navigate to a specific page number, line number, footnote, table, comment, or other object."
      Dim ToolTipFooterTitle As String = "Press F1 for more help."
						Dim MenuItemData As New MenuItemData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = TooTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = ToolTipFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/GoTo_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "G" _
						}
      _dataCollection(Str) = MenuItemData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property


  Public Shared ReadOnly Property Replace() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Replace"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ReplaceToolTipTitle As String = "Replace (Ctrl+H)"
      Dim ReplaceToolTipDescription As String = "Replace text in the document."

						Dim ReplaceButtonData As New ButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Replace_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ReplaceToolTipTitle, _
							Key .ToolTipDescription = ReplaceToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "R" _
						}
      _dataCollection(Str) = ReplaceButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property [Select]() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Select"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Select"
      Dim ToolTipDescription As String = "Select text or objects in the document." & vbLf & vbCr & vbLf & vbCr & "Use Select Object to allow you to select objects that had been positioned behind the text."

						Dim SelectMenuButtonData As New MenuButtonData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Select_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "SL" _
						}
      _dataCollection(Str) = SelectMenuButtonData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SelectAll() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Select _All"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Select All (Ctrl+A)"
      Dim ToolTipDescription As String = "Select all items"

						Dim Data As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Select_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "A" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SelectObjects() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Select _Objects"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Select Objects"
      Dim ToolTipDescription As String = "Select rectangular regions of ink strokes, shapes and text"

						Dim Data As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/Select_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "O" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SelectAllText() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "_Select All Text With Similar Formatting (No Data)"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New MenuItemData() With { _
							Key .Label = Str, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "S" _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property SelectionPane() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Selection _Pane..."

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = "Selection Pane"
      Dim ToolTipDescription As String = "Show the Selection Pane to help select individual objects and to change their order and visibility."

						Dim Data As New MenuItemData() With { _
							Key .Label = Str, _
							Key .SmallImage = New Uri("/RibbonWindowSample;component/Images/SelectionPane_16x16.png", UriKind.Relative), _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "P", _
							Key .IsCheckable = True _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

#End Region

#Region "Insert Table Group Model"

  Public Shared ReadOnly Property Table() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Table"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim ToolTipTitle As String = Str
      Dim ToolTipDescription As String = "Insert or draw a table into the document."

						_dataCollection(Str) = New MenuButtonData() With { _
							Key .LargeImage = New Uri("/RibbonWindowSample;component/Images/Table_32x32.png", UriKind.Relative), _
							Key .Label = Str, _
							Key .ToolTipTitle = ToolTipTitle, _
							Key .ToolTipDescription = ToolTipDescription, _
							Key .ToolTipFooterTitle = HelpFooterTitle, _
							Key .ToolTipFooterImage = New Uri("/RibbonWindowSample;component/Images/Help_16x16.png", UriKind.Relative), _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .KeyTip = "T" _
						}
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property TableGallery() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Table Gallery"

     If Not _dataCollection.ContainsKey(Str) Then
      Dim galleryCategoryData As New GalleryCategoryData(Of RowColumnCount)()
      For i As Integer = 1 To 8
       For j As Integer = 1 To 10
								galleryCategoryData.GalleryItemDataCollection.Add(New RowColumnCount() With { _
									Key .RowCount = i, _
									Key .ColumnCount = j _
								})
       Next
      Next

						Dim galleryData As New GalleryData(Of RowColumnCount)() With { _
							Key .Command = New PreviewDelegateCommand(Of RowColumnCount)(AddressOf InsertTable, AddressOf CanInsertTable, AddressOf PreviewInsertTable, AddressOf CancelPreviewInsertTable), _
							Key .SelectedItem = galleryCategoryData.GalleryItemDataCollection(0) _
						}

      galleryData.CategoryDataCollection.Add(galleryCategoryData)
      _dataCollection(Str) = galleryData
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Private Shared Sub InsertTable(parameter As RowColumnCount)
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.InsertTable()
   End If
  End Sub

  Private Shared Function CanInsertTable(parameter As RowColumnCount) As Boolean
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    Return wordControl__1.CanInsertTable()
   End If

   Return False
  End Function

  Private Shared Sub PreviewInsertTable(parameter As RowColumnCount)
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.PreviewInsertTable(parameter)
   End If
  End Sub

  Private Shared Sub CancelPreviewInsertTable()
   Dim wordControl__1 As UserControlWord = WordControl
   If wordControl__1 IsNot Nothing Then
    wordControl__1.CancelPreviewInsertTable()
   End If
  End Sub

  Public Shared ReadOnly Property DesignTab() As TabData
   Get
    SyncLock _lockObject
     Dim Str As String = "Design"

     If Not _miscData.ContainsKey(Str) Then
						Dim designTabData As New TabData() With { _
							Key .Header = Str, _
							Key .ContextualTabGroupHeader = "Table Tools" _
						}
      _miscData(Str) = designTabData
     End If

     Return TryCast(_miscData(Str), TabData)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LayoutTab() As TabData
   Get
    SyncLock _lockObject
     Dim Str As String = "Layout"

     If Not _miscData.ContainsKey(Str) Then
						Dim layoutTabData As New TabData() With { _
							Key .Header = Str, _
							Key .ContextualTabGroupHeader = "Table Tools" _
						}
      _miscData(Str) = layoutTabData
     End If

     Return TryCast(_miscData(Str), TabData)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property TableTabGroup() As ContextualTabGroupData
   Get
    SyncLock _lockObject
     Dim Str As String = "Table Tools"

     If Not _miscData.ContainsKey(Str) Then
						Dim tableData As New ContextualTabGroupData() With { _
							Key .Header = Str _
						}
      tableData.TabDataCollection.Add(DesignTab)
      tableData.TabDataCollection.Add(LayoutTab)

      _miscData(Str) = tableData
     End If

     Return TryCast(_miscData(Str), ContextualTabGroupData)
    End SyncLock
   End Get
  End Property


  Public Shared ReadOnly Property HeaderRow() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Header Row"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New CheckBoxData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = Str, _
							Key .ToolTipDescription = "Display special formatting for the first row of the table.", _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsChecked = True _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property FirstColumn() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "First Column"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New CheckBoxData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = Str, _
							Key .ToolTipDescription = "Display special formatting for the first column of the table.", _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsChecked = True _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property TotalRow() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Total Row"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New CheckBoxData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = Str, _
							Key .ToolTipDescription = "Display special formatting for the last row of the table.", _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property LastColumn() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Last Column"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New CheckBoxData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = Str, _
							Key .ToolTipDescription = "Display special formatting for the last column of the table.", _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property BandedRows() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Banded Rows"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New CheckBoxData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = Str, _
							Key .ToolTipDescription = "Display banded rows.", _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute), _
							Key .IsChecked = True _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property

  Public Shared ReadOnly Property BandedColumns() As ControlData
   Get
    SyncLock _lockObject
     Dim Str As String = "Banded Columns"

     If Not _dataCollection.ContainsKey(Str) Then
						Dim Data As New CheckBoxData() With { _
							Key .Label = Str, _
							Key .ToolTipTitle = Str, _
							Key .ToolTipDescription = "Display banded columns.", _
							Key .Command = New DelegateCommand(AddressOf DefaultExecuted, AddressOf DefaultCanExecute) _
						}
      _dataCollection(Str) = Data
     End If

     Return _dataCollection(Str)
    End SyncLock
   End Get
  End Property
#End Region

  Private Shared Sub DefaultExecuted()
  End Sub

  Private Shared Function DefaultCanExecute() As Boolean
   Return True
  End Function

  Private Shared ReadOnly Property WordControl() As UserControlWord
   Get
    If Application.Current.Properties.Contains("WordControlRef") Then
     Dim wordControlRef As WeakReference = TryCast(Application.Current.Properties("WordControlRef"), WeakReference)
     If wordControlRef IsNot Nothing Then
      Return TryCast(wordControlRef.Target, UserControlWord)
     End If
    End If
    Return Nothing
   End Get
  End Property


#Region "Data"

  Private Const HelpFooterTitle As String = "Press F1 for more help."
  Private Shared _lockObject As New Object()
  Private Shared _dataCollection As New Dictionary(Of String, ControlData)()

  ' Store any data that doesnt inherit from ControlData
  Private Shared _miscData As New Dictionary(Of String, Object)()
  Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
   target = value
   Return value
  End Function

#End Region

 End Class
End Namespace
