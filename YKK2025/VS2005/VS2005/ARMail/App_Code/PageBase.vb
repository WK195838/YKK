Imports Microsoft.VisualBasic

Public Class PageBase
    Inherits System.Web.UI.Page

    Private Sub Page_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
        Me.Theme = "SkinFile"
    End Sub
End Class
