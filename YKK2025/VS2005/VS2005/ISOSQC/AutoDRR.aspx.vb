Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class AutoDRR
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
        Response.Cookies("PGM").Value = "AutoDRR.aspx"   '程式名
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
        SQL = SQL & "[CreateUserName],[ModifyUserName], "
        '
        SQL = SQL & "[HSCode], "
        SQL = SQL & "'http://10.245.1.6/ISOSQC/CombineList.aspx?pUserID='+'" & UserID & "'+'&pHSCode='+HSCode+'&pPuller='+Puller+'&pColor='+Color as HSCodeURL, "
        '
        SQL = SQL & "[Puller],[Color],[ColorName],[BYColorCode],[Buyer],[BuyerName],[Remark], "
        SQL = SQL & "CASE WHEN CREATEUSER='SP_DRR' THEN 'DRR自動' ELSE '合併自動' END AS CREATEUSER, "
        '
        SQL = SQL & "'http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID='+'" & UserID & "'+'&pBuyer='+Buyer+'&pSlider='+Puller+Color+'&pSource=ISIP' as PullerURL, "
        '
        SQL = SQL & "[DTM_YOBI1],[DTM_YOBI2], "
        SQL = SQL & "[IRW_YOBI1],[IRW_YOBI2], "
        SQL = SQL & "[OR_YOBI2],[Yobi1],[Yobi2] "
        SQL = SQL & "FROM V_PullerColorMagHSCode "
        '
        SQL = SQL & "where CREATEUSER in ('SP_DRR','SP_Combine') "
        ' Puller
        If DKPullerCode.Text <> "" Then
            SQL = SQL & "and puller = '" & DKPullerCode.Text & "' "
        End If
        ' Color
        If DKColor.Text <> "" Then
            SQL = SQL & "and Color = '" & DKColor.Text & "' "
        End If
        ' Buyer
        If DKBuyer.Text <> "" Then
            SQL = SQL & "and buyer + buyername like '%" & DKBuyer.Text & "%' "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [Puller]+[Color]+[ColorName]+[BYColorCode]+[DTM_YOBI1]+[DTM_YOBI2] like '%" & DKOther.Text & "%' "
        End If
        ' AtDRR
        If AtDRR.Checked = True Then
            SQL = SQL & "and CREATEUSER in ('SP_DRR') "
        End If
        ' AtCOMBINE
        If AtCombine.Checked = True Then
            SQL = SQL & "and CREATEUSER in ('SP_Combine') "
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
    '**     Recovery
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        '
        If UCase(Trim(e.CommandName)) = "RECOVERY" Then
            Dim SQL As String
            Dim xHSCode, xPuller, xColor As String
            Dim xIndex As Integer = Convert.ToInt32(e.CommandArgument)
            Dim lnkHSCode As HyperLink = GridView1.Rows(xIndex).Cells(3).Controls(0)
            Dim lnkPuller As HyperLink = GridView1.Rows(xIndex).Cells(1).Controls(0)
            '
            xHSCode = UCase(Replace(lnkHSCode.Text, "&nbsp;", ""))
            xPuller = UCase(Replace(lnkPuller.Text, "&nbsp;", ""))
            xColor = UCase(Replace(GridView1.Rows(xIndex).Cells(2).Text, "&nbsp;", ""))
            'If xHSCode <> "" And xPuller <> "" And xColor <> "" Then
            If xHSCode <> "" And xPuller <> "" Then
                SQL = "exec sp_ISOS2Puller_Combine_R '" & xHSCode & "','" & xPuller & "','" & xColor & "' "
                uDataBase.ExecuteNonQuery(SQL)
            Else
                uJavaScript.PopMsg(Me, "  [HSCODE], [Puller] 必須有資料 !! ")
            End If
        End If

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
            tcl(1).Text = "Puller" & "<BR>" & "Code"
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Color" & "<BR>" & "Code"
            tcl(2).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "HS" & "<BR>" & "Code"
            tcl(3).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Color" & "<BR>" & "Name"
            tcl(4).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Buyer" & "<BR>" & "Color"
            tcl(5).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Buyer" & "<BR>" & ""
            tcl(6).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Buyer" & "<BR>" & "Name"
            tcl(7).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Remark" & "<BR>" & ""
            tcl(8).BackColor = Color.Blue

            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "Auto" & "<BR>" & "Create"
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
            e.Row.Cells(4).Font.Bold = True

            e.Row.Cells(6).ForeColor = Color.Red
            e.Row.Cells(7).ForeColor = Color.Red

            e.Row.Cells(9).Font.Bold = True
            e.Row.Cells(9).ForeColor = Color.Red


            ' 多BUYER換行
            For i = 4 To 15
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


    Protected Sub AtDRR_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtDRR.CheckedChanged
        AtCombine.Checked = False
    End Sub

    Protected Sub AtCombine_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCombine.CheckedChanged
        AtDRR.Checked = False

    End Sub

End Class
