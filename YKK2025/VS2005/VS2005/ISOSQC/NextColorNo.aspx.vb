Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class NextColorNo
    Inherits System.Web.UI.Page
    '
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
    Dim oWaves As New WAVES.CommonService
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
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
        Response.Cookies("PGM").Value = "NextColorNo.aspx"  '程式名")
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Find Puller
    '**
    '*****************************************************************
    Protected Sub BFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFind.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "SELECT "
        SQL = SQL & "[Puller],[LastColor],[NextColorName] "
        SQL = SQL & "FROM V_PullerNextColor "
        '
        SQL = SQL & "where Puller <> 'zzzzzz' "
        ' Puller
        If DKPullerCode.Text <> "" Then
            SQL = SQL & "and puller = '" & DKPullerCode.Text & "' "
        End If
        ' Color
        If DKColor.Text <> "" Then
            SQL = SQL & "and ( "
            SQL = SQL & "    LastColor = '" & DKColor.Text & "' "
            SQL = SQL & " or NextColorName = '" & DKColor.Text & "' "
            SQL = SQL & ") "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [Puller]+[LastColor]+[NextColorName] like '%" & DKOther.Text & "%' "
        End If
        '
        SQL = SQL & "Order by puller, len(puller+LastColor) desc, LastColor desc "
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            '
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Puller" & "<BR>" & "Code"
            tcl(0).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Last Color" & "<BR>" & "Code"
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Next Color" & "<BR>" & "Code"
            tcl(2).BackColor = Color.Blue
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 顏色+格式
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            e.Row.Cells(0).ForeColor = Color.Red
            e.Row.Cells(0).Font.Bold = True
            e.Row.Cells(1).Font.Bold = True
            e.Row.Cells(2).Font.Bold = True
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
End Class
