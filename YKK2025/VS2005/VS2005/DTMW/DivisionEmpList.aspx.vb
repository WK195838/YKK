Imports System.Data
Imports System.Data.OleDb
Partial Class DivisionEmpList
    Inherits System.Web.UI.Page
    Dim Empid As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            EmpData() '取得資料
        End If
    End Sub

    Sub EmpData()
        Dim SQL As String
        Dim UserID As String = Request.QueryString("userid")


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " SELECT  com_code,com_name,id,[Name],Hrwdivid,hrwdivname FROM M_EMP WHERE Com_Code in('01','60','65','70','71')  "
        SQL = SQL + "  AND ( Leav_CD <> 'Y'   OR  (Leav_CD = 'Y' AND LeavDate >= '" & Now.Year & "/" & Now.Month & "/1" & "'  )   )"
        SQL = SQL + " AND ( SupportDivName  in ("
        SQL = SQL + "select data  from M_Referp Where cat='1998' and dkey = 'APPLYDIV-" + UserID + "' )"
        SQL = SQL + " or HRWDivName  in (select data  from M_Referp Where cat='1998' and dkey ='APPLYDIV-" + UserID + "' ) )"

        If Empid <> "" Then
            SQL = SQL + " and ( id = '" + Empid + "' or name like '%" + Empid + "%')"
        End If

        SQL = SQL + " order by hrwdivname,id "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "EmpData")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DNo.TextChanged
        Empid = DNo.Text
        EmpData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()


        Dim DEmpID As String = GridView1.SelectedRow.Cells(2).Text



        Dim js As String = ""
        js &= "var obj = window.opener.document.all.D1;"
        js &= "obj.value = '" & DEmpID & "';"
    
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub

 
    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        EmpData()
    End Sub
End Class
