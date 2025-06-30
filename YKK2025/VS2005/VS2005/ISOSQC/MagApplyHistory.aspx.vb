Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class MagApplyHistory
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
        Response.Cookies("PGM").Value = "MagApplyHistory.aspx"   '程式名
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
        Dim xCode As Integer
        '
        xCode = 0
        If DKADDStart.Text = "" And DKADDEnd.Text = "" And DKKey.Text = "" Then xCode = 9010
        If DKADDStart.Text = "" And DKADDEnd.Text <> "" Then xCode = 9020
        If DKADDStart.Text <> "" And DKADDEnd.Text = "" Then xCode = 9030
        If DKADDStart.Text <> "" And DKADDEnd.Text <> "" Then
            If Len(DKADDStart.Text) <> 8 Or Len(DKADDEnd.Text) <> 8 Then xCode = 9040
            If InStr(DKADDStart.Text, "/") > 0 Then xCode = 9050
            If InStr(DKADDEnd.Text, "/") > 0 Then xCode = 9060
            If DKADDEnd.Text < DKADDStart.Text Then xCode = 9070
        End If
        '
        If xCode = 0 Then
            SQL = "SELECT "
            SQL = SQL & "[CreateUserName],[ModifyUserName], convert(varchar, CreateTime, 111) as CreateTimeDesc, "
            SQL = SQL & "[Puller],[Color],[ColorName],[BYColorCode],[Buyer],[BuyerName],[Remark], "
            '
            SQL = SQL & "'http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID='+'" & UserID & "'+'&pBuyer='+Buyer+'&pSlider='+Puller+Color+'&pSource=ISIP' as PullerURL, "
            '
            SQL = SQL & "[DTM_YOBI1],[DTM_YOBI2], "
            SQL = SQL & "[IRW_YOBI1],[IRW_YOBI2], "
            SQL = SQL & "[OR_YOBI2],[Yobi1],[Yobi2] "
            SQL = SQL & "FROM V_PullerColorMag "
            '
            SQL = SQL & "where buyer <> 'ZZZZZZZ' "
            If DKADDStart.Text <> "" Then
                SQL = SQL & "and convert(varchar, CreateTime, 112) >= '" & DKADDStart.Text & "' "
            End If
            If DKADDEnd.Text <> "" Then
                SQL = SQL & "and convert(varchar, CreateTime, 112) <= '" & DKADDEnd.Text & "' "
            End If
            '
            If AtIncludeSystem.Checked = False And AtIncludeAdmin.Checked = False Then
                SQL = SQL & "and CreateUserName <>  '系統自動' "
                SQL = SQL & "and CreateUserName <>  '管理者' "
            Else
                If AtIncludeSystem.Checked = True And AtIncludeAdmin.Checked = True Then
                Else
                    If AtIncludeSystem.Checked = True And AtIncludeAdmin.Checked = False Then
                        SQL = SQL & "and CreateUserName <>  '管理者' "
                    Else
                        If AtIncludeSystem.Checked = False And AtIncludeAdmin.Checked = True Then
                            SQL = SQL & "and CreateUserName <>  '系統自動' "
                        End If
                    End If
                End If
            End If

            If DKKey.Text <> "" Then
                SQL = SQL & "and [Puller]+[Color]+[ColorName]+[BYColorCode]+[DTM_YOBI1]+[DTM_YOBI2] like '%" & DKKey.Text & "%' "
            End If
            '
            SQL = SQL & "Order by puller, len(puller+color) desc, color desc "
            '
            GridView1.Visible = True
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "[" & Str(xCode) & "] 資料篩選內容無法判別，請檢查 !! ")
        End If
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
            tcl(0).Text = "建立者" & "<BR>" & ""
            tcl(0).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "修改者" & "<BR>" & ""
            tcl(1).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "建立日" & "<BR>" & ""
            tcl(2).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Puller" & "<BR>" & "Code"
            tcl(3).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Color" & "<BR>" & "Code"
            tcl(4).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Color" & "<BR>" & "Name"
            tcl(5).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Buyer" & "<BR>" & "Color"
            tcl(6).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Buyer" & "<BR>" & ""
            tcl(7).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Buyer" & "<BR>" & "Name"
            tcl(8).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "Remark" & "<BR>" & ""
            tcl(9).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "Drop" & "<BR>" & ""
            tcl(10).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "Drop" & "<BR>" & "Remark"
            tcl(11).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "M_Buyer" & "<BR>" & ""
            tcl(12).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "M_Buyer" & "<BR>" & "Name"
            tcl(13).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "Tape" & "<BR>" & "Color"
            tcl(14).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(15).Text = "Data" & "<BR>" & "Source(1)"
            tcl(15).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(16).Text = "Data" & "<BR>" & "Source(2)"
            tcl(16).BackColor = Color.Green
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
            e.Row.Cells(2).ForeColor = Color.Red

            e.Row.Cells(4).ForeColor = Color.Red
            e.Row.Cells(5).ForeColor = Color.Red

            e.Row.Cells(8).Font.Bold = True
            e.Row.Cells(9).Font.Bold = True

        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

End Class
