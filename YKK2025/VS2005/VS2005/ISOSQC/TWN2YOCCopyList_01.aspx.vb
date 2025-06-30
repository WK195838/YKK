Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class TWN2YOCCopyList_01
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
        Response.Cookies("PGM").Value = "TWN2YOCCopyList_01.aspx"   '程式名
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
        AtCloseIMGW.Style("left") = -500 & "px"
        AtCloseIMGW.Checked = False
        DRDImage.Visible = False
        '
        AtCloseWIMGW.Style("left") = -500 & "px"
        AtCloseWIMGW.Checked = False
        DWImage.Visible = False
        '
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
        GridView1.Visible = False
        DRDImage.Visible = False
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
        SQL = "select top 300 "
        SQL = SQL & "SliderCode, "
        SQL = SQL & "YOC,Buyer,ActualDate,ActualTarget,Forcast,Remark, "
        '
        SQL = SQL & "CopyNo,CopyDate, "
        SQL = SQL & "CASE WHEN CopyNo<>'' THEN 'http://10.245.1.6/SPDSub/SPD_YKKGroupCopySheet_02.aspx?pFormNo='+ CopyFormNo + '&pFormSno=' + RTRIM(LTRIM(STR(CopyFormSno))) "
        SQL = SQL & "     ELSE '' "
        SQL = SQL & "END AS CopyNoUrl, "
        '
        SQL = SQL & "MapNo,RDMapNo,RDMapDate, "
        SQL = SQL & "CASE WHEN RDMapFormNo='000001' THEN 'http://10.245.1.10/WorkFlow/MapSheet_02.aspx?pFormNo='+ RDMapFormNo + '&pFormSno=' + RTRIM(LTRIM(STR(RDMapFormSno))) "
        SQL = SQL & "     WHEN RDMapFormNo='000002' THEN 'http://10.245.1.10/WorkFlow/MapModSheet_02.aspx?pFormNo='+ RDMapFormNo + '&pFormSno=' + RTRIM(LTRIM(STR(RDMapFormSno))) "
        SQL = SQL & "     else '' "
        SQL = SQL & "END AS RDMapNooUrl, "
        '
        SQL = SQL & "RDNo, RDDate, RDSuppiler,RDSliderGRCode, "
        SQL = SQL & "CASE WHEN RDFormNo='000003' THEN 'http://10.245.1.10/WorkFlow/ManufInSheet_02.aspx?pFormNo='+ RDFormNo + '&pFormSno=' + RTRIM(LTRIM(STR(RDFormSno))) "
        SQL = SQL & "     WHEN RDFormNo='000007' THEN 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo='+ RDFormNo + '&pFormSno=' + RTRIM(LTRIM(STR(RDFormSno))) "
        SQL = SQL & "     else '' "
        SQL = SQL & "END AS RDNooUrl, "
        '
        SQL = SQL & "OR_NO, OR_SalesDate,OR_Item, "
        SQL = SQL & "CASE WHEN OR_NO<>'' THEN 'http://10.245.1.6/WorkFlowSub/PC_INQWingsOrder.aspx?pOrderNo=' + OR_NO + '&pPuller=' "
        SQL = SQL & "     else '' "
        SQL = SQL & "END AS ORNOUrl, "
        '
        SQL = SQL & "Case When OR_NO<>'' then 'IMG' else '' end as OR_IMG, "
        SQL = SQL + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & UserID & "&pBuyer=" & "" & "&pBuyerItem=' + LTrim(RTrim(OR_Item)) + '&pFun=IMG" & "' as ORIMGURL, "
        '
        SQL = SQL + "'' as xx "
        '
        SQL = SQL & "from W_TWN2YOCCopyInf "
        SQL = SQL & "where SliderCode <> 'zzzzzzzz' "
        'Slider
        If DKSliderCode.Text <> "" Then
            SQL = SQL & "and SliderCode+RDSliderGRCode like '%" & DKSliderCode.Text & "%' "
        End If
        'BUYER
        If DKBuyer.Text <> "" Then
            SQL = SQL & "and Buyer like '%" & DKBuyer.Text & "%' "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and  SliderCode+YOC+Buyer+ActualDate+ActualTarget+Forcast+Remark+CopyNo+CopyDate+MapNo+RDMapNo+RDMapDate+RDSliderGRCode "
            SQL = SQL & " Like"
            SQL = SQL & "'%" & DKOther.Text & "%' "
        End If
        '
        If DKSliderCode.Text = "" And DKBuyer.Text = "" And DKOther.Text = "" Then
            SQL = SQL & "Order by CopyDate Desc, MapNo, SliderCode "
        Else
            SQL = SQL & "Order by MapNo, CopyDate Desc, SliderCode "
        End If
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     WINGS IMAGES
    '**
    '*****************************************************************
    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Dim SQL As String
        Dim xItem As String
        '
        xItem = Replace(GridView1.Rows(e.NewEditIndex).Cells(20).Text, "&nbsp;", "")
        If xItem <> "" Then
            SQL = "select top 1 "
            SQL = SQL & "ImagePath As Path "
            SQL = SQL & "from M_RDPullerImage "
            SQL = SQL & "where formno in ('WINGS')  "
            SQL = SQL & "and Yobi1 = '" & xItem & "' "
            SQL = SQL & "Order by Yobi1 desc "
            '
            Dim dt_RDImage As DataTable = uDataBase.GetDataTable(SQL)
            If dt_RDImage.Rows.Count > 0 Then
                '
                DRDImage.ImageUrl = ""
                If File.Exists(dt_RDImage.Rows(0).Item("Path")) Then
                    DRDImage.ImageUrl = dt_RDImage.Rows(0).Item("Path")
                End If
                DWImage.Visible = True
                '
                AtCloseIMGW.Style("left") = -500 & "px"
                AtCloseIMGW.Checked = False
                '
                AtCloseWIMGW.Style("left") = 600 & "px"
                AtCloseWIMGW.Checked = False
                '
            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     DRAWING DOCUMENT
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim SQL, str, cmd, xSlider As String
        '
        xSlider = Replace(GridView1.SelectedRow.Cells(0).Text, "&nbsp;", "")
        '
        SQL = "select RDMapFormNo, MapFile "
        SQL = SQL & "from W_TWN2YOCCopyInf "
        SQL = SQL & "where SliderCode = '" & xSlider & "' "
        SQL = SQL & "  and RDMapFormNo in ('000001','000002') "
        SQL = SQL & "  and MapFile <> '' "
        SQL = SQL & "Order by RDMapFormNo, MapFile "
        '
        Dim dt_Mapfile As DataTable = uDataBase.GetDataTable(SQL)
        If dt_Mapfile.Rows.Count > 0 Then
            '
            If dt_Mapfile.Rows(0).Item("RDMapFormNo") = "000001" Then
                str = "\\10.245.1.10\WorkFlow\Document\000001\" & dt_Mapfile.Rows(0).Item("MapFile")
                If File.Exists(str) Then
                    str = "file://10.245.1.10/WorkFlow/Document/000001/" & dt_Mapfile.Rows(0).Item("MapFile")
                Else
                    str = "\\10.245.1.9\SyncData\inetpub\wwwroot\WorkFlow\Document\000001\" & dt_Mapfile.Rows(0).Item("MapFile")
                    If File.Exists(str) Then
                        str = "file://10.245.1.9/SyncData/inetpub/wwwroot/WorkFlow/Document/000001/" & dt_Mapfile.Rows(0).Item("MapFile")
                    End If
                End If
                cmd = "<script>" + _
                            "window.open('" & str & "','',''); " + _
                      "</script>"
                Response.Write(cmd)
            End If
            '
            If dt_Mapfile.Rows(0).Item("RDMapFormNo") = "000002" Then
                str = "\\10.245.1.10\WorkFlow\Document\000002\" & dt_Mapfile.Rows(0).Item("MapFile")
                If File.Exists(str) Then
                    str = "file://10.245.1.10/WorkFlow/Document/000002/" & dt_Mapfile.Rows(0).Item("MapFile")
                Else
                    str = "\\10.245.1.9\SyncData\inetpub\wwwroot\WorkFlow\Document\000002\" & dt_Mapfile.Rows(0).Item("MapFile")
                    If File.Exists(str) Then
                        str = "file://10.245.1.9/SyncData/inetpub/wwwroot/WorkFlow/Document/000002/" & dt_Mapfile.Rows(0).Item("MapFile")
                    End If
                End If
                cmd = "<script>" + _
                            "window.open('" & str & "','',''); " + _
                      "</script>"
                Response.Write(cmd)
            End If
            '
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     R&D Images
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
        Dim lnk As HyperLink = GridView1.Rows(e.RowIndex).Cells(8).Controls(0)
        '
        SQL = "select top 1 [ImagePath] as Path, [Yobi1] as BackupPath "
        SQL = SQL & "from M_RDPullerImage "
        SQL = SQL & "where No = '" & Replace(lnk.Text, "&nbsp;", "") & "' "
        SQL = SQL & "Order by createtime "
        '
        Dim dt_RDImage As DataTable = uDataBase.GetDataTable(SQL)
        If dt_RDImage.Rows.Count > 0 Then
            '
            DRDImage.ImageUrl = ""
            If File.Exists(dt_RDImage.Rows(0).Item("Path")) Then
                DRDImage.ImageUrl = dt_RDImage.Rows(0).Item("Path")
            Else
                If File.Exists(dt_RDImage.Rows(0).Item("BackupPath")) Then
                    DRDImage.ImageUrl = dt_RDImage.Rows(0).Item("BackupPath")
                End If
            End If
            DRDImage.Visible = True
            '
            AtCloseIMGW.Style("left") = 600 & "px"
            AtCloseIMGW.Checked = False
            '
            AtCloseWIMGW.Style("left") = -500 & "px"
            AtCloseWIMGW.Checked = False
            '
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            ' detail
            '添加新的表头第一行
            '
            '1~8
            i = 0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Slider" & "<BR>" & "Code"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "YOC" & "<BR>" & ""
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Buyer" & "<BR>" & ""
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Date" & "<BR>" & ""
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Target" & "<BR>" & ""
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Drawing" & "<BR>" & "No"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Forcast" & "<BR>" & ""
            tcl(i).BackColor = Color.Blue
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Remark" & "<BR>" & ""
            tcl(i).BackColor = Color.Blue
            '
            '14~22
            '9~17
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Develop" & "<BR>" & "Doc."
            tcl(i).BackColor = Color.Purple
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Date"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Suppiler"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Puller"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Images"
            tcl(i).BackColor = Color.Purple
            '
            '9~13
            '18~22
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Transfer" & "<BR>" & "Doc."
            tcl(i).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Date"
            tcl(i).BackColor = Color.Green
            '
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Drawing" & "<BR>" & "Doc."
            tcl(i).BackColor = Color.Green
            i = i + 1
            '
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Date"
            tcl(i).BackColor = Color.Green
            '
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "" & "<BR>" & "Images"
            tcl(i).BackColor = Color.Green
            '
            'i = i + 1
            'tcl.Add(New TableHeaderCell())
            'tcl(i).Text = "WINGS" & "<BR>" & "OR"
            'tcl(i).BackColor = Color.Green
            'i = i + 1
            'tcl.Add(New TableHeaderCell())
            'tcl(i).Text = "" & "<BR>" & "Date"
            'tcl(i).BackColor = Color.Green
            'i = i + 1
            'tcl.Add(New TableHeaderCell())
            'tcl(i).Text = "" & "<BR>" & "Item"
            'tcl(i).BackColor = Color.Green
            'i = i + 1
            'tcl.Add(New TableHeaderCell())
            'tcl(i).Text = "" & "<BR>" & "Images"
            'tcl(i).BackColor = Color.Green
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "Transfer"
            H3tc1.ColumnSpan = 8
            H3row.Cells.Add(H3tc1)
            '
            Dim H5tc1 As TableCell = New TableCell
            H5tc1.Text = "Develop"
            H5tc1.ColumnSpan = 10
            H3row.Cells.Add(H5tc1)

            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
        '
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
            For i = 0 To 1
                e.Row.Cells(i).ForeColor = Color.Red
            Next
            '
            For i = 3 To 4
                e.Row.Cells(i).Font.Bold = True
            Next
            '
            For i = 18 To 21
                e.Row.Cells(i).Visible = False
            Next
            ' Images lINK
            e.Row.Cells(12).ForeColor = Color.Blue
            e.Row.Cells(17).ForeColor = Color.Blue
            e.Row.Cells(21).ForeColor = Color.Blue
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
    '**(Close RD IMG Windows)
    '**     
    '**
    '*****************************************************************
    Protected Sub AtCloseIMGW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseIMGW.CheckedChanged
        DRDImage.Visible = False
        AtCloseIMGW.Style("left") = -500 & "px"
        AtCloseIMGW.Checked = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Close WINGS IMG Windows)
    '**     
    '**
    '*****************************************************************
    Protected Sub AtCloseWIMGW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseWIMGW.CheckedChanged
        DWImage.Visible = False
        AtCloseWIMGW.Style("left") = -500 & "px"
        AtCloseWIMGW.Checked = False
    End Sub
End Class
