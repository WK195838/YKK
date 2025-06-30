Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_ReplaceItem01
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim wBuyer As String             'Buyer
    Dim wUserID As String            'UserID
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
            ShowData()
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
        Response.Cookies("PGM").Value = "PS_INQ_ReplaceItem01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        wBuyer = Request.QueryString("pBuyer")
        wUserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBUYER.Text = ""
        HBuyerCode.Text = ""
        DKEY.Text = ""
        '
        If wBuyer <> "" Then
            If wBuyer = "000021" Then DBUYER.Text = "TNF"
            HBuyerCode.Text = wBuyer
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '
        If wBuyer <> "" Then
            sql = "SELECT top 300 "
            sql &= "'S' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*

            If wBuyer = "000021" Then sql &= "'http://10.245.0.3/test/FindItemPage.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + D1 as URL "
            'If wBuyer = "000021" Then sql &= "'http://10.245.1.6/WorkFlowSub/FindItemPage.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + D1 as URL "

            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And DataType = '" & "REPLACEITEM" & "' "
            sql &= "And Active = 0 "
            '
            If DKEY.Text <> "" Then sql &= "And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY.Text & "%' "
            '*
            sql &= "Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=CustomerOrder.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'Buyer
        If wBuyer = "000021" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "Buyer Item"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "YKK Item"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "Remark"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 5 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

End Class
