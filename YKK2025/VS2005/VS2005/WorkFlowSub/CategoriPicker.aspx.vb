Imports System.Data.SqlClient
Imports System.Data
Partial Class CategoriPicker
    Inherits System.Web.UI.Page
    Dim UserID As String
    Dim sqlstrcat As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        UserID = Request.QueryString("pUSerID")
        GIVE_BIND()
      
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim Dcat As String = GridView1.SelectedRow.Cells(1).Text
        Dim DName1 As String = GridView1.SelectedRow.Cells(2).Text


        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DCat;"
        js &= "obj.value = '" & Dcat & "';"
        js &= "var obj1 = window.opener.document.all.DName;"
        js &= "obj1.value = '" & DName1 & "';"
        js &= "var obj2 = window.opener.document.all.DNo;"
        js &= "obj2.value = '" & Dcat & "';"
        'js &= "var obj4 = window.opener.document.getElementById('DName');"
        'js &= "obj4.innerText = '" & DName1 & "';"
        js &= "window.close();"
     
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)


    End Sub
    Sub GIVE_BIND()
        Dim sqlStr As String = ""
        Dim sqlStr1 As String = ""
        sqlStr = "Select count(*) from m_referp where cat ='9001' and DKEY ='" & USerID & "'"

        Dim dbcon As New SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("WorkFlow_Con").ToString)
        Dim dbcmd As New SqlCommand(sqlStr, dbcon)

        dbcmd.Connection.Open()
        If dbcmd.ExecuteScalar >= 1 Then
            sqlStr1 = "Select * from m_referp where  cat= '999' and dkey like '%" & DCat.Text & "%'"
        Else
            sqlStr1 = " select substring(dkey,8,4)dkey,data  from m_referp where  cat =  '9002' and dkey like '%" & UserID & "%'"
        End If
        'Response.Write(sqlStr1)
        dbcmd.Connection.Close()
        Dim dbcmd1 As New SqlCommand(sqlStr1, dbcon)

        Dim dbAdpter As New SqlDataAdapter(dbcmd1)
        Dim ds As New Data.DataSet
        dbAdpter.Fill(ds)

        GridView1.DataSource = ds
        GridView1.DataBind()
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSerch.Click
        GIVE_BIND()

    End Sub

End Class
