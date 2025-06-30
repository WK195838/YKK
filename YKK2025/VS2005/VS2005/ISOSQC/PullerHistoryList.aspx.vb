Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PullerHistoryList
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
        Response.Cookies("PGM").Value = "PullerHistoryList.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID

        AtCloseIMGW.Style("left") = -500 & "px"
        AtCloseIMGW.Checked = False
        DRDImage.Visible = False
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
        If DKPullerCode.Text <> "" Then
            '
            SQL = "select "
            SQL = SQL & "HSCODE, [Puller], [Color], [ColorName], [BYColorCode], [Buyer], [BuyerName], [DTM_YOBI1], "
            SQL = SQL & "CASE WHEN [DTM_YOBI1]='X' THEN LTRIM(RTRIM([DTM_YOBI2])) + '|' + [Remark] ELSE [Remark] END AS [Remark], "
            '
            SQL = SQL & "[DTM], [IRW], [ORDERS], "
            '
            SQL = SQL & "CASE WHEN RD_YOBI1>'1' THEN RD+' *' ELSE RD END AS RD, "
            SQL = SQL & "CASE WHEN RD_YOBI1>'1' THEN 'http://10.245.1.6/ISOSQC/ManySupplierList.aspx?pUserID=' + '" & UserID & "' + '&pPuller=' + Puller + '&pColor=' + Color + '&pFun=MANYRD' ELSE '' END AS RDMUrl, "
            SQL = SQL & "CASE WHEN EDX_YOBI1>'1' THEN EDX+' *' ELSE EDX END AS EDX, "
            'SQL = SQL & "'' END AS EDXMUrl, "
            SQL = SQL & "CASE WHEN EDX_YOBI1>'1' THEN 'http://10.245.1.6/ISOSQC/ManySupplierList.aspx?pUserID=' + '" & UserID & "' + '&pPuller=' + Puller + '&pColor=' + Color + '&pFun=' + [EDX_NO] ELSE '' END AS EDXMUrl, "
            '
            SQL = SQL & "Case When [RD_NO]<>'' then [RD_NO] else '' end as RD_NO, "
            SQL = SQL & "CASE WHEN RD_NO<>'' THEN 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo='+ RD_FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(RD_FORMSNO))) ELSE '' END AS RDUrl, "
            SQL = SQL & "[RD_AppDate], [RD_Supplier], "
            '
            SQL = SQL & "Case When [EDX_NO]<>'' then [EDX_NO] else '' end as EDX_NO, "
            SQL = SQL & "Case When [EDX_NO]<>'' and [EDX_NO]<>'QC5Y' then 'http://10.245.1.6/ISOSQC/QASheet_02.aspx?pFormNo=008002&pFormSno=' + ltrim(rtrim(str(EDX_formsno))) "
            SQL = SQL & "     else 'http://10.245.1.6/ISOSQC/QCEDXList.aspx?pPuller=' + Puller + '&pCOLOR=' + Color "
            SQL = SQL & "end as EDXUrl, "
            SQL = SQL & "[EDX_AppDate], [EDX_Supplier], "
            '
            SQL = SQL & "CASE WHEN ISNULL((SELECT TOP 1 [ImagePath] FROM [M_RDPullerImage] WHERE [SliderGRCode] =B_PullerColor.PULLER + B_PullerColor.COLOR AND FORMNO IN ('008002')), '') = '' THEN '' "
            SQL = SQL & "     ELSE 'IMG' "
            SQL = SQL & "END as EDX_IMG, "
            SQL = SQL & "'http://10.245.1.6/ISOSQC/Http2File.aspx?pUserID=' + '" & UserID & "' + '&pNo=' + [EDX_NO] as EDXIMGUrl, "
            '
            SQL = SQL & "Case When [IRW_NO]<>'' then [IRW_NO] else '' end as IRW_NO, "
            SQL = SQL & "'http://10.245.1.6/IRW/ItemRegisterSheet_02.aspx?pFormNo=001151&pFormSno=' + ltrim(rtrim(str(IRW_formsno))) as IRWUrl, "
            SQL = SQL & "[IRW_AppDate], "
            '
            SQL = SQL & "Case When [OR_NO]<>'' then [OR_NO] else '' end as OR_NO, "
            SQL = SQL & "'http://10.245.1.6/WorkFlowSub/PC_INQWingsOrder.aspx?pOrderNo=' + OR_NO as ORUrl, "
            SQL = SQL & "[OR_CfmDate], "

            SQL = SQL & "Case When [OR_NO]<>'' then 'IMG' else '' end as OR_IMG, "
            SQL = SQL + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & UserID & "&pBuyer=" & "" & "&pBuyerItem=' + LTrim(RTrim(OR_YOBI1)) + '&pFun=IMG" & "' as ORIMGURL, "
            '
            SQL = SQL & "[OR_YOBI2] AS TapeColor "
            '
            SQL = SQL & "FROM B_PullerColor "
            '
            SQL = SQL & "where puller = '" & DKPullerCode.Text & "' "
            '
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
            SQL = SQL & "Order by puller, len(puller+color) desc, color desc, hscode desc "
            '
            ' MsgBox(SQL)
            GridView1.Visible = True
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "  [Puller Code] 必須輸入 !! ")
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
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "" & "<BR>" & ""
            tcl(0).BackColor = Color.Black
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "" & "<BR>" & ""
            tcl(1).BackColor = Color.Black
            'BUYER PULLER
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Puller" & "<BR>" & "Code"
            tcl(2).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Color" & "<BR>" & "Code"
            tcl(3).BackColor = Color.Blue
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
            tcl(9).Text = "Drop" & "<BR>" & ""
            tcl(9).BackColor = Color.Blue
            'STATUS
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "RD" & "<BR>" & ""
            tcl(10).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "DTM" & "<BR>" & ""
            tcl(11).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "EDX" & "<BR>" & ""
            tcl(12).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "IRW" & "<BR>" & ""
            tcl(13).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "ORDERS" & "<BR>" & ""
            tcl(14).BackColor = Color.Purple
            'LINK
            tcl.Add(New TableHeaderCell())
            tcl(15).Text = "RD" & "<BR>" & "NO."
            tcl(15).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(16).Text = "RD" & "<BR>" & "App.d"
            tcl(16).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(17).Text = "RD" & "<BR>" & "Supplier"
            tcl(17).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(18).Text = "RD" & "<BR>" & "Images"
            tcl(18).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(19).Text = "EDX" & "<BR>" & "NO."
            tcl(19).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(20).Text = "EDX" & "<BR>" & "App.d"
            tcl(20).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(21).Text = "EDX" & "<BR>" & "Supplier"
            tcl(21).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(22).Text = "EDX" & "<BR>" & "Images"
            tcl(22).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(23).Text = "IRW" & "<BR>" & "NO."
            tcl(23).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(24).Text = "IRW" & "<BR>" & "App.d"
            tcl(24).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(25).Text = "OR" & "<BR>" & "NO."
            tcl(25).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(26).Text = "OR" & "<BR>" & "Sales.d"
            tcl(26).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(27).Text = "OR" & "<BR>" & "Images"
            tcl(27).BackColor = Color.Green
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "Puller Color"
            H3tc1.ColumnSpan = 10
            H3row.Cells.Add(H3tc1)
            '
            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "Progress Status"
            H4tc1.ColumnSpan = 5
            H3row.Cells.Add(H4tc1)
            '
            Dim H5tc1 As TableCell = New TableCell
            H5tc1.Text = "Link"
            H5tc1.ColumnSpan = 13
            H3row.Cells.Add(H5tc1)

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
        Dim i, j As Integer
        Dim sts As String
        Dim BuyerInf As String()
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(0).BackColor = Color.Black
            e.Row.Cells(0).ForeColor = Color.White

            e.Row.Cells(2).ForeColor = Color.Red
            e.Row.Cells(3).ForeColor = Color.Red
            '
            e.Row.Cells(4).Font.Bold = True
            e.Row.Cells(5).Font.Bold = True
            ' 多BUYER對應
            For i = 6 To 6
                If DKBuyer.Text <> "" And e.Row.Cells(i).Text <> DKBuyer.Text And _
                   InStr(e.Row.Cells(i).Text, DKBuyer.Text) > 0 And InStr(e.Row.Cells(i).Text, "|") > 0 Then
                    ' 6:buyer
                    BuyerInf = e.Row.Cells(i).Text.Split("|")
                    For j = 0 To BuyerInf.Length - 1
                        If BuyerInf(j) = DKBuyer.Text Then Exit For
                    Next
                    e.Row.Cells(6).Text = BuyerInf(j)
                    ' 4:colorname
                    BuyerInf = e.Row.Cells(4).Text.Split("|")
                    e.Row.Cells(4).Text = BuyerInf(j)
                    ' 5:buyercolor
                    BuyerInf = e.Row.Cells(5).Text.Split("|")
                    e.Row.Cells(5).Text = BuyerInf(j)
                    ' 7:buyername
                    BuyerInf = e.Row.Cells(7).Text.Split("|")
                    e.Row.Cells(7).Text = BuyerInf(j)
                    ' 8:remark
                    BuyerInf = e.Row.Cells(8).Text.Split("|")
                    e.Row.Cells(8).Text = BuyerInf(j)
                    ' 28:TapeColor
                    BuyerInf = e.Row.Cells(28).Text.Split("|")
                    If BuyerInf(j) <> "" Then
                        If e.Row.Cells(8).Text = "&nbsp;" Then
                            e.Row.Cells(8).Text = "Tape Color=[" & BuyerInf(j) & "]"
                        Else
                            e.Row.Cells(8).Text = e.Row.Cells(8).Text & " Tape Color=[" & BuyerInf(j) & "]"
                        End If
                    End If
                Else
                    If Replace(e.Row.Cells(28).Text, "|", "") <> "" And e.Row.Cells(28).Text <> "&nbsp;" Then
                        If e.Row.Cells(8).Text = "&nbsp;" Then
                            e.Row.Cells(8).Text = "Tape Color=[" & e.Row.Cells(28).Text & "]"
                        Else
                            e.Row.Cells(8).Text = e.Row.Cells(8).Text & "|" & "Tape Color=[" & e.Row.Cells(28).Text & "]"
                        End If
                    End If
                End If
            Next
            ' 多BUYER換行
            For i = 4 To 8
                e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, "|", "<br>")
            Next
            ' Drop Status
            For i = 9 To 9
                If e.Row.Cells(i).Text = "X" Then
                    e.Row.Cells(i).BackColor = Color.Red
                    e.Row.Cells(i).ForeColor = Color.White
                    e.Row.Cells(i).Font.Bold = True
                End If
            Next
            ' Progress Status
            For i = 10 To 16
                If InStr(e.Row.Cells(i).Text, "WIP") > 0 Then
                    e.Row.Cells(i).BackColor = Color.Cyan
                    e.Row.Cells(i).ForeColor = Color.Black
                    e.Row.Cells(i).Font.Bold = True
                End If
                '
                If i = 10 Then
                    Dim lnk As HyperLink = e.Row.Cells(i).Controls(0)
                    If InStr(lnk.Text, "WIP") > 0 Then
                        e.Row.Cells(i).BackColor = Color.Cyan
                        e.Row.Cells(i).ForeColor = Color.Black
                        e.Row.Cells(i).Font.Bold = True
                    End If
                End If
                '
                If i = 12 Then
                    Dim lnk As HyperLink = e.Row.Cells(i).Controls(0)
                    If InStr(lnk.Text, "WIP") > 0 Then
                        e.Row.Cells(i).BackColor = Color.Cyan
                        e.Row.Cells(i).ForeColor = Color.Black
                        e.Row.Cells(i).Font.Bold = True
                    End If
                End If
            Next
            ' RD & EDX Images lINK
            For i = 18 To 18
                If e.Row.Cells(i - 1).Text = "&nbsp;" Then e.Row.Cells(i).Text = ""
                e.Row.Cells(i).ForeColor = Color.Blue
            Next
            ' TAPE COLOR (OR_YOBI2)
            e.Row.Cells(28).Visible = False
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
    '**     R&D Images
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
        Dim lnk As HyperLink = GridView1.Rows(e.RowIndex).Cells(15).Controls(0)
        '
        SQL = "select top 1 [ImagePath] as Path "
        SQL = SQL & "from M_RDPullerImage "
        SQL = SQL & "where No = '" & Replace(lnk.Text, "&nbsp;", "") & "' "
        SQL = SQL & "Order by createtime "
        Dim dt_RDImage As DataTable = uDataBase.GetDataTable(SQL)
        If dt_RDImage.Rows.Count > 0 Then
            DRDImage.ImageUrl = dt_RDImage.Rows(0).Item("Path")
            DRDImage.Visible = True
            '
            AtCloseIMGW.Style("left") = 1064 & "px"
            AtCloseIMGW.Checked = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Close IMG Windows)
    '**     
    '**
    '*****************************************************************
    Protected Sub AtCloseIMGW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCloseIMGW.CheckedChanged
        AtCloseIMGW.Style("left") = -500 & "px"
        AtCloseIMGW.Checked = False
    End Sub
End Class
