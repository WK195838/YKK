Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SPDStatusRebuild
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
        Response.Cookies("PGM").Value = "StatusRebuild.aspx"   '程式名
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
        DKPullerCode.Text = UCase(DKPullerCode.Text)
        DKColor.Text = UCase(DKColor.Text)
        '
        If DKPullerCode.Text <> "" Then
            SQL = "SELECT "
            SQL = SQL & "[CreateUserName],[ModifyUserName], "
            SQL = SQL & "[RD],[DTM],[EDX],[IRW],[ORDERS], "
            SQL = SQL & "[Puller],[Color],[ColorName],[BYColorCode],[Buyer],[BuyerName],[Remark], "
            '
            SQL = SQL & "'http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID='+'" & UserID & "'+'&pBuyer='+Buyer+'&pSlider='+Puller+Color+'&pSource=ISIP' as PullerURL, "
            '
            SQL = SQL & "[DTM_YOBI1],[DTM_YOBI2], "
            SQL = SQL & "[IRW_YOBI1],[IRW_YOBI2], "
            SQL = SQL & "[OR_YOBI2],[Yobi1],[Yobi2] "
            SQL = SQL & "FROM V_PullerColorMag "
            '
            SQL = SQL & "where CREATEUSER not in ('SP_DRR') "
            ' Puller
            If DKPullerCode.Text <> "" Then
                SQL = SQL & "and puller = '" & DKPullerCode.Text & "' "
            End If
            ' Color
            If DKColor.Text <> "" Then
                SQL = SQL & "and Color Like '%" & DKColor.Text & "%' "
            End If
            '
            SQL = SQL & "Order by puller, len(puller+color) desc, color desc "
            '
            GridView1.Visible = True
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "  [Puller] 必須輸入 !! ")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Rebuild
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        '
        If UCase(Trim(e.CommandName)) = "REFRESH" Then
            Dim SQL As String
            Dim xPuller, xColor As String
            Dim xIndex As Integer = Convert.ToInt32(e.CommandArgument)
            Dim lnk As HyperLink = GridView1.Rows(xIndex).Cells(6).Controls(0)
            '
            xPuller = UCase(Replace(lnk.Text, "&nbsp;", ""))
            xColor = UCase(Replace(GridView1.Rows(xIndex).Cells(7).Text, "&nbsp;", ""))
            If xPuller <> "" And xColor <> "" Then
                SQL = "exec sp_SPD2EDXIRW_Manual '" & xPuller & "','" & xColor & "' "
                uDataBase.ExecuteNonQuery(SQL)
            Else
                uJavaScript.PopMsg(Me, "  [Puller], [Color] 必須有資料 !! ")
            End If
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
            tcl(0).Text = "" & "<BR>" & ""
            tcl(0).BackColor = Color.White
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "RD" & "<BR>" & ""
            tcl(1).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "DTM" & "<BR>" & ""
            tcl(2).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "EDX" & "<BR>" & ""
            tcl(3).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "IRW" & "<BR>" & ""
            tcl(4).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "ORDERS" & "<BR>" & ""
            tcl(5).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Puller" & "<BR>" & "Code"
            tcl(6).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Color" & "<BR>" & "Code"
            tcl(7).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Color" & "<BR>" & "Name"
            tcl(8).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "Buyer" & "<BR>" & "Color"
            tcl(9).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "Buyer" & "<BR>" & ""
            tcl(10).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "Buyer" & "<BR>" & "Name"
            tcl(11).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "Remark" & "<BR>" & ""
            tcl(12).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "Drop" & "<BR>" & ""
            tcl(13).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "Drop" & "<BR>" & "Remark"
            tcl(14).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(15).Text = "M_Buyer" & "<BR>" & ""
            tcl(15).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(16).Text = "M_Buyer" & "<BR>" & "Name"
            tcl(16).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(17).Text = "Tape" & "<BR>" & "Color"
            tcl(17).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(18).Text = "Data" & "<BR>" & "Source(1)"
            tcl(18).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(19).Text = "Data" & "<BR>" & "Source(2)"
            tcl(19).BackColor = Color.Green
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
        Dim jsScript As String
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            ' Drop Status
            For i = 1 To 5
                If e.Row.Cells(i).Text = "WIP" Then
                    e.Row.Cells(i).BackColor = Color.Cyan
                    e.Row.Cells(i).ForeColor = Color.Black
                    e.Row.Cells(i).Font.Bold = True
                End If
            Next
            e.Row.Cells(7).ForeColor = Color.Red
            e.Row.Cells(8).ForeColor = Color.Red

            e.Row.Cells(11).Font.Bold = True
            e.Row.Cells(12).Font.Bold = True

            ' 多BUYER換行
            For i = 9 To 19
                e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, "|", "<br>")
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
