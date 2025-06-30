Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class CategoriPicker
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim GroupCode As String         'GroupCode

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then
            GetData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                  '設定逾時時間
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DGridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim Sql As String = ""
        Sql &= "SELECT Cat, Dkey, Data "
        Sql &= "FROM M_Referp "
        Sql &= "Where Cat = '999' "
        Sql &= "order by Cat, DKey, Data "

        Dim dt_Referp As DataTable = uDataBase.GetDataTable(Sql)
        If dt_Referp.Rows.Count > 0 Then
            DGridView1.Visible = True
            DGridView1.DataSource = dt_Referp
            DGridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DGridView1)
    '**     換頁處理
    '**
    '*****************************************************************
    Protected Sub DGridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGridView1.PageIndexChanged
        GetData()
    End Sub

    Protected Sub DGridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles DGridView1.PageIndexChanging
        DGridView1.PageIndex = e.NewPageIndex
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BGo_Click)
    '**     查詢處理
    '**
    '*****************************************************************
    Protected Sub BGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGo.Click
        Dim Sql As String = ""
        Sql &= "SELECT Cat, Dkey, Data "
        Sql &= "FROM M_Referp "
        Sql &= "Where Cat = '999' "
        If DDKey.Text <> "" Then
            Sql &= "And   (DKey Like '%" & DDKey.Text & "%' OR Data Like '%" & DDKey.Text & "%') "
        End If
        Sql &= "order by Cat, DKey, Data "

        Dim dt_Referp As DataTable = uDataBase.GetDataTable(Sql)
        If dt_Referp.Rows.Count > 0 Then
            DGridView1.Visible = True
            DGridView1.DataSource = dt_Referp
            DGridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DGridView1)
    '**     選取資料後回傳父視窗
    '**
    '*****************************************************************
    Protected Sub DGridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles DGridView1.RowCommand
        If e.CommandName = "Pick" Then
            SelectReturn(e.CommandArgument)
        End If

    End Sub

    Sub SelectReturn(ByVal cat As String)
        Dim uJScript As New Utility.JScript
        Dim sControlId As String() = Request.QueryString("pControlID").Split(",")
        Dim sAttribute As String() = {"value", "value"}
        Dim sValue As String() = {cat, cat}
        uJScript.RegJavaScript(Me, "picker", uJScript.ReturnValue(sControlId, sAttribute, sValue))
    End Sub

End Class
