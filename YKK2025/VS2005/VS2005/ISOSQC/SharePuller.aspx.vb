Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SharePuller
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
        Response.Cookies("PGM").Value = "SharePuller.aspx"   '程式名
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
        DKOther.Text = UCase(DKOther.Text)
        '
        SQL = "SELECT TOP 150 "
        SQL = SQL & "case when active='1' then 'CLOSE' else 'OPEN' end as Sts, "
        SQL = SQL & "case when active='1' then '停止使用' else '使用中' end as StsDesc, "
        SQL = SQL & "[Puller],[Color],YOBI1,YOBI2,  convert(varchar, CreateTime, 111) as CreateTime "
        SQL = SQL & "FROM M_ExcludePullerColor "
        '
        SQL = SQL & "where active <> 9 "
        ' Puller
        If DKPullerCode.Text <> "" Then
            SQL = SQL & "and puller = '" & DKPullerCode.Text & "' "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [Puller]+[Color]+YOBI1+YOBI2 like '%" & DKOther.Text & "%' "
        End If
        ' AtOPEN
        If AtOPEN.Checked = True Then
            SQL = SQL & "and ACTIVE = 0 "
        End If
        ' AtDTM
        If AtCLOSE.Checked = True Then
            SQL = SQL & "and ACTIVE = 1 "
        End If
        '
        SQL = SQL & "Order by puller, len(puller+color) desc, color desc "
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
            tcl(0).Text = "Status" & "<BR>" & ""
            tcl(0).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "" & "<BR>" & ""
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Puller" & "<BR>" & ""
            tcl(2).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Color" & "<BR>" & ""
            tcl(3).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Yobi1" & "<BR>" & ""
            tcl(4).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Yobi2" & "<BR>" & ""
            tcl(5).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Create" & "<BR>" & "Time"
            tcl(6).BackColor = Color.Green
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
            For i = 0 To 6
                If i = 0 Then
                    If InStr(e.Row.Cells(i).Text, "CLOSE") > 0 Then
                        e.Row.Cells(i).BackColor = Color.Red
                        e.Row.Cells(i).ForeColor = Color.White
                    End If
                End If
                '
                e.Row.Cells(2).ForeColor = Color.Red
                e.Row.Cells(2).Font.Bold = True
                '
                e.Row.Cells(6).Font.Bold = True
            Next

        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub

    Protected Sub AtOPEN_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtOPEN.CheckedChanged
        AtCLOSE.Checked = False
    End Sub

    Protected Sub AtCLOSE_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCLOSE.CheckedChanged
        AtOPEN.Checked = False
    End Sub

End Class
