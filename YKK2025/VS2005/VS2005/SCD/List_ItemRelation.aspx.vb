Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class List_ItemRelation
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDate As String           '現在日期
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPopupFunction()                          '設定彈出視窗事件
        If Not IsPostBack Then                      'PostBack
            SetSearchField()                        '設定搜尋欄位
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
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "List_ItemRelation.aspx"                                    '程式名
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")                      '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")                            '工程代碼
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDate = Now.ToString("yyyy/MM/dd")                  '現在日期
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetSearchField)
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchField()
        Dim sql As String = ""
        Dim dtFieldData As DataTable
        'SIZE
        DSize.Items.Clear()
        sql = "Select SIZENO From M_ItemRelation Group by SIZENO Order by SIZENO "
        dtFieldData = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dtFieldData.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtFieldData.Rows(i).Item("SIZENO")
            ListItem1.Value = dtFieldData.Rows(i).Item("SIZENO")
            DSize.Items.Add(ListItem1)
        Next
        'CHAIN
        DChain.Items.Clear()
        sql = "Select ITEM From M_ItemRelation Group by ITEM Order by ITEM "
        dtFieldData = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dtFieldData.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtFieldData.Rows(i).Item("ITEM")
            ListItem1.Value = dtFieldData.Rows(i).Item("ITEM")
            DChain.Items.Add(ListItem1)
        Next
        '供應商
        DSup.Items.Clear()
        'DSup.Items.Add("ALL")
        sql = "Select COL From M_ItemRelation Group by COL Order by COL "
        dtFieldData = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dtFieldData.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtFieldData.Rows(i).Item("COL")
            ListItem1.Value = dtFieldData.Rows(i).Item("COL")
            DSup.Items.Add(ListItem1)
        Next
        '色番
        DColor.Text = ""
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BOK_Click)
    '**     OK按鈕按下事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     篩選資料處理
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String = ""
        sql = "SELECT SIZENO, ITEM, COL, Colno, CODE "
        sql &= "FROM M_ItemRelation "
        sql &= "WHERE STS = '0' "
        sql &= "  AND SIZENO =  '" + DSize.SelectedValue + "' "
        sql &= "  AND ITEM   =  '" + DChain.SelectedValue + "' "
        sql &= "  AND COL    =  '" + DSup.SelectedValue + "' "
        '色番
        If DColor.Text <> "" Then
            sql &= " AND Colno LIKE '%" + DColor.Text + "%' "
        End If
        sql &= "ORDER BY SIZENO, ITEM, COL, Colno, CODE "
        '
        Dim dtItem As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtItem
        GridView1.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Gridview1_PageIndexChanging)
    '**     換頁處理
    '**
    '*****************************************************************
    Protected Sub Gridview1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub
    Protected Sub Gridview1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        ShowData()
    End Sub
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
    End Sub

    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Response.AppendHeader("Content-Disposition", "attachment;filename=SCD_ItemRelationList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub

End Class
