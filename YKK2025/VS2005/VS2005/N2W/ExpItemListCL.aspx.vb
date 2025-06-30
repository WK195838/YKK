Imports System.Data
Imports System.Data.OleDb

Partial Class ExpItemListCL
    Inherits System.Web.UI.Page

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim InputData As String
    Dim Depo As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetExpItem()
            GetData()
        End If
    End Sub

    Sub GetExpItem()
        Dim SQL As String
        Dim i As Integer

        Depo = Request.QueryString("field")

        SQL = " select* from M_referp  where dkey = 'salesdepo'"
        SQL = SQL & " and data like '%" & Depo & "%'"
        Dim DBDepo As DataTable = uDataBase.GetDataTable(SQL)
        If DBDepo.Rows.Count > 0 Then
            SQL = " Select ExpCat from  M_ExpItemListCL "  '販管
            SQL = SQL + "  where left(acid,1)<>'5' group by ExpCat "
            SQL = SQL + " order by ExpCat "
        Else
            SQL = " Select ExpCat from  M_ExpItemListCL "  '製造
            SQL = SQL + "  where left(acid,1)<>'6' group by ExpCat "
            SQL = SQL + " order by ExpCat "
        End If

        DExpCat.Items.Clear()
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter2.Rows.Count > 0 Then
            For i = 0 To DBAdapter2.Rows.Count - 1
                DExpCat.Items.Add(DBAdapter2.Rows(i).Item("ExpCat"))
            Next
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        'Dim UserID As String = Request.QueryString("userid")


        Depo = Request.QueryString("field")

        SQL = " select* from M_referp  where dkey = 'salesdepo'"
        SQL = SQL & " and data like '%" & Depo & "%'"
        Dim DBDepo As DataTable = uDataBase.GetDataTable(SQL)

        If DBDepo.Rows.Count > 0 Then
            SQL = " Select * from  M_ExpItemListCL "  '販管
            SQL = SQL + "  where left(acid,1)<>'5' and  ExpCat = '" & DExpCat.SelectedValue & "' "
            If InputData <> "" Then
                SQL = SQL + " and ( ExpItem like '%" + InputData + "%')"
            End If
            SQL = SQL + " order by ExpCat "
        Else
            SQL = " Select * from  M_ExpItemListCL "  '製造
            SQL = SQL + "  where left(acid,1)<>'6' and  ExpCat = '" & DExpCat.SelectedValue & "' "
            If InputData <> "" Then
                SQL = SQL + " and ( ExpItem like '%" + InputData + "%')"
            End If
            SQL = SQL + " order by ExpCat "

        End If
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

        GridView1.DataSource = DBAdapter2
        GridView1.DataBind()


    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = DData.Text
        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
        '
        Dim Exp As String = GridView1.SelectedRow.Cells(1).Text & "--" & GridView1.SelectedRow.Cells(4).Text
        Dim ACID As String = GridView1.SelectedRow.Cells(3).Text
        '
        Dim js As String = ""
        js &= "window.opener.document.all.DExpItem1.value = '" & Exp & "';"
        js &= "window.opener.document.all.DACID.value = '" & ACID & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"
        '
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)
    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        e.Row.Cells(3).Visible = False
    End Sub
End Class
