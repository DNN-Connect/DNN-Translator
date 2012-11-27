Imports System.Text.RegularExpressions
Imports DotNetNuke.Translator.Common

Namespace ViewModel
 Public Class SearchViewModel
  Inherits ResourceCollectionViewModel

  Private _basePath As String = ""
  Public Property SearchString As String = ""
  Public Property TooManyResults As Boolean = False

  Public Sub New(parameters As Common.ParameterList)
   MyBase.New(CType(parameters.ParentWindow, MainWindowViewModel))
   Me.ShowFileColumn = True
   _basePath = MainWindow.ProjectSettings.Location
   Me.SearchString = parameters.Params(1)
   Me.ID = String.Format("{0}_{1}_{2}", parameters.Params(0), parameters.Params(1), TargetLocale.Name)
   Me.DisplayName = "Search Results"
   AddResources(parameters.Params(0))
  End Sub

  Private Sub AddResources(path As String)

   If Me.ResourceKeys.Count > 100 Then
    TooManyResults = True
    Exit Sub
   End If

   Try
    Dim directory As String = (_basePath & path).ToLower
    For Each resFile As Common.ResourceFile In MainWindow.ProjectSettings.CurrentSnapShot.ResourceFiles.Values
     If resFile.FilePath.ToLower.StartsWith(directory) Then
      For Each key As String In resFile.Resources.Keys
       'Dim contents As String = System.Web.HttpUtility.HtmlDecode(resFile.Resources(key).Value)
       Dim contents As String = resFile.Resources(key).Value
       If contents.IndexOf(SearchString, StringComparison.InvariantCultureIgnoreCase) > -1 Then
        AddKey(New ResourceKeyViewModel(resFile.Resources(key), Common.Globals.GetTranslation(_basePath, resFile.fileKey, key, TargetLocale.Name)))
       End If
      Next
     End If
    Next

   Catch ex As Exception

   End Try

  End Sub

 End Class
End Namespace
