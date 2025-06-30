Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class MPMProcessesRecord
    Inherits System.Web.UI.Page


 
    Protected Sub DNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.TextChanged
        If DNo.Text <> "" Then
            DNo.Text = DNo.Text + Chr(13)
        End If

    End Sub
End Class

