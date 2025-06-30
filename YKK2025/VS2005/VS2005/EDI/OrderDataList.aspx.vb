Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class OrderDataList
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim wUserID As String           'UserID
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
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then                      'PostBack
            SetSearchItem()
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "OrderDataList.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")   '現在日期時間
        wUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選條件
    '**
    '*****************************************************************
    Sub SetSearchItem()
        Dim sql As String
        '
        DCustomer.Items.Clear()
        sql = "Select Customer From W_EDIData "
        sql &= "Group by Customer "
        sql &= "Order by Customer "
        Dim dt_EDI As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dt_EDI.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt_EDI.Rows(i).Item("Customer")
            ListItem1.Value = dt_EDI.Rows(i).Item("Customer")
            DCustomer.Items.Add(ListItem1)
        Next
        '
        DDate.Text = NowDate
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     設定參數
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        Dim sql As String
        '
        sql = "SELECT Top 30 "
        sql &= "Customer, Buyer, DataDate, DataTime, "
        sql &= "'EDI檔(Excel)' As EDIDesc, EDI, "
        sql &= "'EDI檔-全欄位(Excel)' As EDIAllDesc, EDIAll, "
        sql &= "'客戶原始檔' As EDIOriDesc, EDIOri "
        sql = sql + "FROM W_EDIData "
        sql = sql + "Where Customer = '" & DCustomer.Text & "' "
        If DBuyer.Text <> "ALL" Then
            sql = sql + "  And Buyer = '" & DBuyer.Text & "' "
        End If
        If DDate.Text <> "" Then
            sql = sql + "  And DataDate <= '" & DDate.Text & "' "
        End If
        sql = sql + "Order by DataDate Desc, DataTime Desc, Buyer "
        GridView1.DataSource = uDataBase.GetDataTable(sql)
        GridView1.DataBind()
    End Sub

End Class
