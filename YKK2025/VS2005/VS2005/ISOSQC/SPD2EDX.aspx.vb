Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SPD2EDX
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
        Response.Cookies("PGM").Value = "SPD2EDX.aspx"   '程式名
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
        SQL = SQL & "'X' AS EDX, "
        SQL = SQL & "[R_Puller],[R_Color],[R_RDNO],[R_CompDateTime],[R_Supplier], "
        SQL = SQL & "'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo='+ R_FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(R_FORMSNO))) AS R_RDNOURL, "
        ' --
        SQL = SQL & "[Puller],[Color],[ColorName],[BYColorCode],[Buyer],[BuyerName],[Remark], "
        SQL = SQL & "'http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID='+'" & UserID & "'+'&pBuyer='+Buyer+'&pSlider='+Puller+Color+'&pSource=ISIP' as PullerURL, "
        '
        SQL = SQL & "[DTM_YOBI1],[DTM_YOBI2], "
        SQL = SQL & "[IRW_YOBI1],[IRW_YOBI2], "
        SQL = SQL & "[OR_YOBI2],[Yobi1],[Yobi2] "
        SQL = SQL & "FROM V_SPDEDXInf "
        '
        SQL = SQL & "where R_Puller <> '' "
        ' Puller
        If DKPullerCode.Text <> "" Then
            SQL = SQL & "and R_Puller = '" & DKPullerCode.Text & "' "
        End If
        ' Color
        If DKColor.Text <> "" Then
            SQL = SQL & "and R_Color = '" & DKColor.Text & "' "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [R_Puller]+[R_Color]+[R_RDNO]+[R_CompDateTime]+[R_Supplier]+[Puller]+[Color]+[ColorName]+[BYColorCode]+[DTM_YOBI1]+[DTM_YOBI2] like '%" & DKOther.Text & "%' "
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
            tcl(0).Text = "EDX" & "<BR>" & ""
            tcl(0).BackColor = Color.White
            tcl(0).ForeColor = Color.Black
            '
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "R&D" & "<BR>" & "Puller"
            tcl(1).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "R&D" & "<BR>" & "Color"
            tcl(2).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "R&D" & "<BR>" & "No."
            tcl(3).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "R&D" & "<BR>" & "Comp.Date"
            tcl(4).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "R&D" & "<BR>" & "Supplier"
            tcl(5).BackColor = Color.Black
            '
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "ISIP" & "<BR>" & "Puller"
            tcl(6).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "ISIP" & "<BR>" & "Color"
            tcl(7).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "ISIP" & "<BR>" & "ColorName"
            tcl(8).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "ISIP" & "<BR>" & "BuyerColor"
            tcl(9).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "ISIP" & "<BR>" & "Buyer"
            tcl(10).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "ISIP" & "<BR>" & "BuyerName"
            tcl(11).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "ISIP" & "<BR>" & "Remark"
            tcl(12).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "ISIP" & "<BR>" & "Drop"
            tcl(13).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "ISIP" & "<BR>" & "DropRemark"
            tcl(14).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(15).Text = "ISIP" & "<BR>" & "M_Buyer"
            tcl(15).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(16).Text = "ISIP" & "<BR>" & "M_BuyerName"
            tcl(16).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(17).Text = "ISIP" & "<BR>" & "TapeColor"
            tcl(17).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(18).Text = "ISIP" & "<BR>" & "DataSource(1)"
            tcl(18).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(19).Text = "ISIP" & "<BR>" & "DataSource(2)"
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
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            e.Row.Cells(1).ForeColor = Color.Red
            e.Row.Cells(2).Font.Bold = True

            e.Row.Cells(7).ForeColor = Color.Red
            e.Row.Cells(8).Font.Bold = True

            e.Row.Cells(10).ForeColor = Color.Red
            e.Row.Cells(11).ForeColor = Color.Red

            For i = 0 To 0
                If e.Row.Cells(i).Text = "X" Then
                    e.Row.Cells(i).BackColor = Color.Red
                    e.Row.Cells(i).ForeColor = Color.White
                    e.Row.Cells(i).Font.Bold = True
                End If
            Next


            ' 多BUYER換行
            For i = 8 To 19
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
