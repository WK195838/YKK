Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_CustomerGroup
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
        Sql &= "SELECT *, "
        Sql &= "'INQ_CustomerGroupHistory.aspx?' + "
        Sql &= "'pGroupCode=' + REPLACE(GroupCode,'&','%26') + "
        Sql &= "'&pUserID=' + '" & UserID & "' "
        Sql &= "As URL "
        Sql &= "FROM M_CustomerGroup "
        Sql &= "Where GroupCode <> '' "

        If Not String.IsNullOrEmpty(DCust.Text.Trim) Then
            Sql &= "and (GroupCode like '%" & DCust.Text.Trim & "%' or GroupName like '%" & DCust.Text.Trim & "%') "
        End If
        If Not String.IsNullOrEmpty(DSalesName.Text.Trim) Then
            Sql &= "and SalesName like '%" & DSalesName.Text.Trim & "%' "
        End If
        If Not String.IsNullOrEmpty(DSalesCode.Text.Trim) Then
            Sql &= "and SalesCode like '%" & DSalesCode.Text.Trim & "%' "
        End If
        Sql &= "order by GroupCode, GroupName, SalesCode "

        Dim dt_CustomerGroup As DataTable = uDataBase.GetDataTable(Sql)
        If dt_CustomerGroup.Rows.Count > 0 Then
            DGridView1.Visible = True
            DGridView1.DataSource = dt_CustomerGroup
            DGridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSearch_Click)
    '**     搜尋資料
    '**
    '*****************************************************************
    Protected Sub BSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSearch.Click
        GetData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub DGridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles DGridView1.RowDataBound
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            'Dim Hyper1 As New HyperLink
            'Dim URL As String = "INQ_CustomerGroupHistory.aspx?" + _
            '                    "pGroupCode=" + DataBinder.Eval(e.Row.DataItem, "GroupCode") + _
            '                    "&pUserID=" + UserID


            'Hyper1.Text = DataBinder.Eval(e.Row.DataItem, "GroupCode")
            'Hyper1.NavigateUrl = URL
            'Hyper1.Target = "_blank"
            'e.Row.Cells(0).Controls.Add(Hyper1)
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub


End Class