Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_MailCustomerGroupHistory
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
        GroupCode = Request.QueryString("pGroupCode")       'GroupCode
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
        Sql &= "SELECT Top 100 A.*, "
        Sql &= "(SELECT GroupName From M_CustomerGroup Where GroupCode = A.GroupCode) As GroupName, "
        Sql &= "MPath + RFile As WFile, "
        Sql &= "MHttpPath + RFile As HFile, "
        Sql &= "Replace(MHttpPath + RFile, '#','%23') As URL, "
        Sql &= "Case isnull(MHttpPath + RFile,'NULL') When 'NULL' Then '' Else '....' End As RAttachFile "
        Sql &= "FROM T_ProcHistory A "
        Sql &= "Where A.GroupCode = '" & GroupCode & "' "
        Sql &= "order by A.Unique_ID Desc "

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
    '**(GridView)
    '**     設定下載Excel
    '**
    '*****************************************************************
    Protected Sub DGridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles DGridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
        End If
    End Sub

End Class
