Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class OrderAmount
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
        Response.Cookies("PGM").Value = "OrderAmount.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID

        AtCloseW.Style("left") = -500 & "px"
        AtCloseW.Checked = False
        GridView2.Visible = False
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
        SQL = SQL & "[Puller],[Color],[ColorName],[BYColorCode],[Buyer],[BuyerName],[Remark], "
        '
        SQL = SQL & "'http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID='+'" & UserID & "'+'&pBuyer='+Buyer+'&pSlider='+Puller+Color+'&pSource=ISIP' as PullerURL, "
        '
        SQL = SQL & "[DTM_YOBI1],[DTM_YOBI2], "
        SQL = SQL & "[IRW_YOBI1],[IRW_YOBI2], "
        SQL = SQL & "[OR_YOBI2],[Yobi1],[Yobi2], "
        '
        SQL = SQL & "[Qty],[Amt],[Qty1Y],[Amt1Y],[Qty2Y],[Amt2Y],[Qty3Y],[Amt3Y],[Qty4Y],[Amt4Y],[Qty5Y],[Amt5Y] "
        '
        SQL = SQL & "FROM V_PullerColorMag10YAmount "
        '
        SQL = SQL & "where Puller <> '' "
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
            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            '
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Puller" & "<BR>" & "Code"
            tcl(0).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Color" & "<BR>" & "Code"
            tcl(1).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Color" & "<BR>" & "Name"
            tcl(2).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Buyer" & "<BR>" & "Color"
            tcl(3).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Buyer" & "<BR>" & ""
            tcl(4).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Buyer" & "<BR>" & "Name"
            tcl(5).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Remark" & "<BR>" & ""
            tcl(6).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Drop" & "<BR>" & ""
            tcl(7).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Drop" & "<BR>" & "Remark"
            tcl(8).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "M_Buyer" & "<BR>" & ""
            tcl(9).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "M_Buyer" & "<BR>" & "Name"
            tcl(10).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "Tape" & "<BR>" & "Color"
            tcl(11).BackColor = Color.Black
            '
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "" & "<BR>" & ""
            tcl(12).BackColor = Color.Black

            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "QTY" & "<BR>" & ""
            tcl(13).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "AMT" & "<BR>" & ""
            tcl(14).BackColor = Color.Green

            tcl.Add(New TableHeaderCell())
            tcl(15).Text = "QTY" & "<BR>" & ""
            tcl(15).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(16).Text = "AMT" & "<BR>" & ""
            tcl(16).BackColor = Color.Blue

            tcl.Add(New TableHeaderCell())
            tcl(17).Text = "QTY" & "<BR>" & ""
            tcl(17).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(18).Text = "AMT" & "<BR>" & ""
            tcl(18).BackColor = Color.Green

            tcl.Add(New TableHeaderCell())
            tcl(19).Text = "QTY" & "<BR>" & ""
            tcl(19).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(20).Text = "AMT" & "<BR>" & ""
            tcl(20).BackColor = Color.Blue

            tcl.Add(New TableHeaderCell())
            tcl(21).Text = "QTY" & "<BR>" & ""
            tcl(21).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(22).Text = "AMT" & "<BR>" & ""
            tcl(22).BackColor = Color.Green

            tcl.Add(New TableHeaderCell())
            tcl(23).Text = "QTY" & "<BR>" & ""
            tcl(23).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(24).Text = "AMT" & "<BR>" & ""
            tcl(24).BackColor = Color.Blue
            '-----------------------------------------
            ' 表頭-N-1行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "過去5年販賣實績"
            H3tc1.ColumnSpan = 13
            H3tc1.BackColor = Color.Red
            H3tc1.ForeColor = Color.White
            H3tc1.Font.Bold = True
            H3tc1.HorizontalAlign = HorizontalAlign.Left
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "N" & "<BR>" & "This Year"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Green
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "N-1" & "<BR>" & "Year"
            H3tc3.ColumnSpan = 2
            H3tc3.BackColor = Color.Blue
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "N-2" & "<BR>" & "Year"
            H3tc4.ColumnSpan = 2
            H3tc4.BackColor = Color.Green
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "N-3" & "<BR>" & "Year"
            H3tc5.ColumnSpan = 2
            H3tc5.BackColor = Color.Blue
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.Text = "N-4" & "<BR>" & "Year"
            H3tc6.ColumnSpan = 2
            H3tc6.BackColor = Color.Green
            H3row.Cells.Add(H3tc6)

            Dim H3tc7 As TableCell = New TableCell
            H3tc7.Text = "N-5" & "<BR>" & "Year"
            H3tc7.ColumnSpan = 2
            H3tc7.BackColor = Color.Blue
            H3row.Cells.Add(H3tc7)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)

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
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            For i = 0 To 24
                Select Case i
                    Case 12
                        e.Row.Cells(i).ForeColor = Color.Red
                    Case 13, 14, 17, 18, 21, 22
                        e.Row.Cells(i).ForeColor = Color.Green
                    Case Else
                        e.Row.Cells(i).ForeColor = Color.Blue

                End Select
                e.Row.Cells(i).Font.Bold = True
            Next
            e.Row.Cells(25).Visible = False
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Dim SQL As String

        SQL = "SELECT "
        SQL &= "[Slider],[Month],[Custmer],[CustmerName],[Buyer],[BuyerName], "
        SQL &= "[Quantity],[Amount],[Puller],[Color],[YY] "
        '
        SQL = SQL & "FROM W_SALES_DATA_10Y_YYMM "
        SQL = SQL & "where puller = '" & Replace(GridView1.Rows(e.NewEditIndex).Cells(25).Text, "&nbsp;", "") & "' "
        SQL = SQL & "  and color = '" & Replace(GridView1.Rows(e.NewEditIndex).Cells(1).Text, "&nbsp;", "") & "' "
        SQL = SQL & "Order By Puller,Color, Month Desc, Custmer, buyer  "
        '
        GridView2.Visible = True
        GridView2.DataSource = uDataBase.GetDataTable(SQL)
        GridView2.DataBind()
        '
        AtCloseW.Style("left") = 900 & "px"
        AtCloseW.Checked = False
        '
    End Sub

    Protected Sub AtCloseW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseW.CheckedChanged
        GridView2.Visible = False
        AtCloseW.Style("left") = -500 & "px"
        AtCloseW.Checked = False
    End Sub

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated

        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            '
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Puller" & "<BR>" & "Code"
            tcl(0).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Color" & "<BR>" & "Code"
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Year" & "<BR>" & "Month"
            tcl(2).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Slider" & "<BR>" & "Code"
            tcl(3).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Customer" & "<BR>" & ""
            tcl(4).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "" & "<BR>" & ""
            tcl(5).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Buyer" & "<BR>" & ""
            tcl(6).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "" & "<BR>" & ""
            tcl(7).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "QTY" & "<BR>" & ""
            tcl(8).BackColor = Color.Red
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "AMT" & "<BR>" & ""
            tcl(9).BackColor = Color.Red
            '
            gv.Controls(0).Controls.AddAt(0, H3row)

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

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            For i = 0 To 9
                Select Case i
                    Case 0, 1
                        e.Row.Cells(i).Font.Bold = True
                    Case 8, 9
                        e.Row.Cells(i).ForeColor = Color.Red
                    Case Else
                End Select
            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
End Class
