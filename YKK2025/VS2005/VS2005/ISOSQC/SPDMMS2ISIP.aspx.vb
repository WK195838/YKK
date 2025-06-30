Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SPDMMS2ISIP
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
        Response.Cookies("PGM").Value = "ISIP2RD.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID

        '動作按鈕設定
        AtCloseRDW.Style("left") = -500 & "px"
        AtCloseRDW.Checked = False
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
        SQL = SQL & "[PULLER], [RDCount], [NoMMSCount], [YesMMSCount] "
        SQL = SQL & "FROM V_Puller2SPDMMS "
        '
        SQL = SQL & "where Puller <> '' "
        ' Puller
        If DKPullerCode.Text <> "" Then
            SQL = SQL & "and puller = '" & DKPullerCode.Text & "' "
        End If
        ' AtYesMMS
        If AtYesMMS.Checked = True Then
            SQL = SQL & "and RDCount>0 and NoMMSCount=0 and YesMMSCount>0 "
        End If
        ' AtYesNoMMS
        If AtYesNoMMS.Checked = True Then
            SQL = SQL & "and RDCount>0 and NoMMSCount>0 and YesMMSCount>0 "
        End If
        ' AtNoMMS
        If AtNoMMS.Checked = True Then
            SQL = SQL & "and RDCount>0 and NoMMSCount>0 and YesMMSCount=0 "
        End If
        ' AtNoSPD
        If AtNoSPD.Checked = True Then
            SQL = SQL & "and RDCount=0 and NoMMSCount=0 and YesMMSCount=0 "
        End If
        '
        SQL = SQL & "Order by puller "
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
            tcl(0).Text = "" & "<BR>" & ""
            tcl(0).BackColor = Color.White

            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Puller" & "<BR>" & ""
            tcl(1).BackColor = Color.Red

            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "RD" & "<BR>" & "Count"
            tcl(2).BackColor = Color.Black

            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "未模廢" & "<BR>" & "Count"
            tcl(3).BackColor = Color.Black

            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "已模廢" & "<BR>" & "Count"
            tcl(4).BackColor = Color.Black
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
            For i = 0 To 4
                e.Row.Cells(i).Font.Bold = True
            Next
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub

    Protected Sub AtYesMMS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtYesMMS.CheckedChanged
        AtYesNoMMS.Checked = False
        AtNoMMS.Checked = False
        AtNoSPD.Checked = False
        AtALL.Checked = False
    End Sub

    Protected Sub AtYesNoMMS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtYesNoMMS.CheckedChanged
        AtYesMMS.Checked = False
        AtNoMMS.Checked = False
        AtNoSPD.Checked = False
        AtALL.Checked = False
    End Sub

    Protected Sub AtNoMMS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtNoMMS.CheckedChanged
        AtYesMMS.Checked = False
        AtYesNoMMS.Checked = False
        AtNoSPD.Checked = False
        AtALL.Checked = False
    End Sub

    Protected Sub AtNoSPD_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtNoSPD.CheckedChanged
        AtYesMMS.Checked = False
        AtNoMMS.Checked = False
        AtYesNoMMS.Checked = False
        AtALL.Checked = False
    End Sub

    Protected Sub AtALL_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtALL.CheckedChanged
        AtYesMMS.Checked = False
        AtNoMMS.Checked = False
        AtYesNoMMS.Checked = False
        AtNoSPD.Checked = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     R&D Information
    '**
    '*****************************************************************
    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Dim SQL As String

        SQL = "SELECT "

        SQL &= "CASE WHEN MMS=1 THEN '模廢' ELSE '' END AS MMSDESC, "
        SQL &= "Puller, SliderCode, Spec, SUPPLIER, Sts AS Status, "

        Sql &= "CASE WHEN RDNO<>'' THEN RDNO ELSE '' END AS RDNOM, "
        Sql &= "CASE WHEN RDNO<>'' THEN 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FORMSNO))) ELSE '' END AS RDNOUrl, "

        Sql &= "CASE WHEN OPContact=1 THEN 'L' ELSE '' END AS OPContactM, "
        Sql &= "CASE WHEN OPContact=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FORMSNO))) + '&pFun=OP' ELSE '' END AS OPContactL, "

        Sql &= "CASE WHEN Contact=1 THEN 'L' ELSE '' END AS ContactM, "
        Sql &= "CASE WHEN Contact=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pFun=CT' ELSE '' END AS ContactL, "

        Sql &= "CASE WHEN SliderDetail=1 THEN 'L' ELSE '' END AS SliderDetailM, "
        Sql &= "CASE WHEN SliderDetail=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pCode=' + PULLER+COLOR + '&pFun=SD' ELSE '' END AS SliderDetailL, "

        Sql &= "CASE WHEN Surface=1 THEN 'L' ELSE '' END AS SurfaceM, "
        Sql &= "CASE WHEN Surface=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pFun=SF' ELSE '' END AS SurfaceL, "

        Sql &= "CASE WHEN ColorAppend=1 THEN 'L' ELSE '' END AS ColorAppendM, "
        Sql &= "CASE WHEN ColorAppend=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pFun=CA' ELSE '' END AS ColorAppendL, "

        Sql &= "CASE WHEN YKKGroupCopy=1 THEN 'L' ELSE '' END AS YKKGroupCopyM, "
        Sql &= "CASE WHEN YKKGroupCopy=1 THEN 'http://10.245.1.6//SPDSub/YKKGroupCopyList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) ELSE '' END AS YKKGroupCopyL "
        '
        SQL = SQL & "FROM V_PullerColor2SPD "
        '
        If InStr("/5N18/6N18/7N18/", "/" & Replace(GridView1.Rows(e.NewEditIndex).Cells(1).Text, "&nbsp;", "") & "/") > 0 Then
            SQL = SQL & "WHERE puller LIKE '_N18' "
        Else
            If InStr("/N5N18/N6N18/N7N18/", "/" & Replace(GridView1.Rows(e.NewEditIndex).Cells(1).Text, "&nbsp;", "") & "/") > 0 Then
                SQL = SQL & "WHERE puller LIKE 'N_N18' "
            Else
                SQL = SQL & "where puller = '" & Replace(GridView1.Rows(e.NewEditIndex).Cells(1).Text, "&nbsp;", "") & "' "
            End If
        End If
        '
        GridView2.Visible = True
        GridView2.DataSource = uDataBase.GetDataTable(Sql)
        GridView2.DataBind()
        '
        AtCloseRDW.Style("left") = 424 & "px"
        AtCloseRDW.Checked = False
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Close RD Windows)
    '**     
    '**
    '*****************************************************************
    Protected Sub AtCloseRDW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseRDW.CheckedChanged
        GridView2.Visible = False
        AtCloseRDW.Style("left") = -500 & "px"
        AtCloseRDW.Checked = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView2 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        '
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            '
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "" & "<BR>" & ""
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Puller" & "<BR>" & "Code"
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Status" & "<BR>" & ""
            tcl(2).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "R&D" & "<BR>" & "No"
            tcl(3).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Slider" & "<BR>" & "Code"
            tcl(4).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Size-Family-Body" & "<BR>" & ""
            tcl(5).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Supplier" & "<BR>" & ""
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "　型別" & "<BR>" & "　胴體"
            tcl(7).BackColor = Color.Red
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "　業務" & "<BR>" & "　連絡"
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "　拉頭" & "<BR>" & "　細目"
            tcl(9).BackColor = Color.Red
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "　表面" & "<BR>" & "　處理"
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "　外注　" & "<BR>" & "　色番"
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "　圖面" & "<BR>" & "　複製"
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If e.Row.Cells(0).Text = "模廢" Then
                e.Row.Cells(0).BackColor = Color.Black
                e.Row.Cells(0).ForeColor = Color.White
                'e.Row.Cells(0).Font.Bold = True
            End If
            e.Row.Cells(1).ForeColor = Color.Red
            e.Row.Cells(2).Font.Bold = True
            e.Row.Cells(6).ForeColor = Color.Red

            e.Row.Cells(4).Font.Bold = True
            e.Row.Cells(5).Font.Bold = True

            e.Row.Cells(7).Font.Bold = True
            e.Row.Cells(8).Font.Bold = True
            e.Row.Cells(9).Font.Bold = True
            e.Row.Cells(10).Font.Bold = True
            e.Row.Cells(11).Font.Bold = True
            e.Row.Cells(12).Font.Bold = True
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
End Class
