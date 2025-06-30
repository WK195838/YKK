Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing

Partial Class SPSActionHistory
    Inherits System.Web.UI.Page


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '
    Dim NowDateTime As String       '現在日期時間
    Dim xUserID, xLogID, xBuyer, xAction
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
            DataList("ALL", 1)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
        '
        xLogID = Request.QueryString("pLogID")
        xBuyer = Request.QueryString("pBuyer")
        xUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetSearchItem)
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchItem()
        DCustBuyer.Text = xBuyer
        DLogID.Text = xLogID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     資料抽出
    '**
    '*****************************************************************
    Sub DataList(ByVal pCat As String, ByVal pFlag As Integer)
        Dim Sql As String
        '
        Dim OleDbConnection1 As New OleDbConnection
        Dim ds1 As New DataSet
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '
        Sql = "SELECT "
        '
        Sql = Sql + "LogID, Buyer, "
        Sql = Sql + "Case Action "
        Sql = Sql + "When 'SPS_ActionPlan' Then 'ActionPlan' "
        Sql = Sql + "When 'SPS_LocalStockPlan' Then 'LocalStockPlan' "
        Sql = Sql + "When 'SPS_ForcastPlan' Then 'ForcastPlan' "
        Sql = Sql + "Else 'Material' End As ActionDesc, "
        '
        Sql = Sql + "Convert(VARCHAR(20), RunTime, 120) As RunTimeDesc, "
        Sql = Sql + "Case Error When 0 Then '正常' Else '異常' End As ErrorDesc, "
        Sql = Sql + "D1, D2, D3, D4, D5 "
        Sql = Sql + "From L_ActionHistory "
        Sql = Sql + "Where LogID = '" + DLogID.Text + "' "
        Sql = Sql + "  And Buyer = '" + DCustBuyer.Text + "' "
        '
        If pCat = "999" Then
            Sql = Sql + "  And D1 = '999' "
        End If
        '
        If pFlag = 1 Then
            Sql = Sql + "  And Error = " + CStr(pFlag) + " "
        End If
        Sql = Sql + "Order by Unique_ID "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, OleDbConnection1)
        DBAdapter1.Fill(ds1, "ActionLog")
        '
        GridView1.DataSource = ds1
        GridView1.DataBind()
        '
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BError)
    '**     按鈕(BError)
    '**
    '*****************************************************************
    Protected Sub BError_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BError.Click
        DataList("ALL", 1)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BAll)
    '**     按鈕(BAll)
    '**
    '*****************************************************************
    Protected Sub BAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BAll.Click
        DataList("ALL", 9)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     顏色處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim xBackColor As Boolean = False
            '--異常--------------------------------------------
            If e.Row.Cells(2).Text.ToString = "異常" Then
                e.Row.ForeColor = Color.Blue
                e.Row.BackColor = Color.LightPink
            End If
        End If
    End Sub
End Class
