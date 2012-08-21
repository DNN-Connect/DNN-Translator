Namespace Services.Packing
 Public Class LanguagePack

  Public Property Manifest As PackManifest

  Public Sub New(root As String, start As String, locale As String, ownerEmail As String, ownerName As String, ownerOrganization As String, ownerUrl As String, license As String, objectName As String, objectManifest As String)
   Me.Manifest = New PackManifest(locale, ownerEmail, ownerName, ownerOrganization, ownerUrl, license, objectName, objectManifest)

  End Sub

  Public Sub Save(path As String)

  End Sub

 End Class
End Namespace
