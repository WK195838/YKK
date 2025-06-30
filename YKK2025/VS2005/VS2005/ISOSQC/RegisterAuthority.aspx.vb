Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class RegisterAuthority
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
        Response.Cookies("PGM").Value = "RegisterAuthority.aspx"   '程式名
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
        SQL = SQL & "[Cat],[UserName],[UserID],[DepName],[BuyerPullerStr] "
        SQL = SQL & "FROM V_NewPullerColorUser "
        '
        SQL = SQL & "where Cat <> 'zzzzz' "
        ' id / name
        If DKUserID.Text <> "" Then
            SQL = SQL & "and (UserID like '%" & DKUserID.Text & "%' or UserName Like '%" & DKUserID.Text & "%') "
        End If
        ' Buyer
        If DKBuyer.Text <> "" Then
            SQL = SQL & "and BuyerPullerStr like '%" & DKBuyer.Text & "%' "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [Cat]+[UserName]+[UserID]+[DepName]+[BuyerPullerStr] like '%" & DKOther.Text & "%' "
        End If
        '
        SQL = SQL & "Order by cat, UserName, BuyerPullerStr "
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
            tcl(0).Text = "Cat." & "<BR>" & ""
            tcl(0).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "UserName" & "<BR>" & ""
            tcl(1).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "UserID" & "<BR>" & ""
            tcl(2).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Division" & "<BR>" & ""
            tcl(3).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "[BuyerPullerStr]" & "<BR>" & ""
            tcl(4).BackColor = Color.Blue
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
            For i = 0 To 8
                e.Row.Cells(0).Font.Bold = True
                '
                e.Row.Cells(2).ForeColor = Color.Red
                e.Row.Cells(2).Font.Bold = True
                e.Row.Cells(4).ForeColor = Color.Red
                e.Row.Cells(4).Font.Bold = True

                ' 多BUYER換行
                If i = 4 Then
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, "/", "<br>")
                End If
            Next

        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
End Class
