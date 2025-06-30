Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class PC_ReleaseNewcolor
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
    Dim pSize As String             'Size
    Dim pFamily As String           'Family
    Dim pSlider As String           'Slider / Puller
    Dim pBuyer As String            'Buyer  / Color
    Dim pSource As String           'Source
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
        Response.Cookies("PGM").Value = "PC_ReleaseNewcolor.aspx"   '程式名
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

        AtCloseIMGW.Style("left") = -500 & "px"
        AtCloseIMGW.Checked = False
        DRDImage.Visible = False

        AtCloseEDXW.Style("left") = -500 & "px"
        AtCloseEDXW.Checked = False
        DEDXImage.Visible = False
        '
        '動作按鈕設定
        DADVREPORTFile.Style("left") = -500 & "px"
        DADVREPORTFile.Text = "\\10.245.0.192\program$\ISOS\DataPrepare\" + "INQPULLERForISIP_" & UserID & ".xlsm"
        BADVReport.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('ADVREPORTExcel')}"
        '
        DADVSALESFile.Style("left") = -500 & "px"
        DADVSALESFile.Text = "\\10.245.0.192\program$\ISOS\DataPrepare\" + "INQSALESForISIP_" & UserID & ".xlsm"
        BADVSales.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('ADVSALESExcel')}"
        '
        DPULLERLISTFile.Style("left") = -500 & "px"
        DPULLERLISTFile.Text = "\\10.245.0.192\program$\ISOS\DataPrepare\" + "INQPULLERLISTForISIP_" & UserID & ".xlsm"
        BPullerList.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('PULLERLISTExcel')}"
        '
        'admin
        BAdmin.Visible = False
        If InStr("IT003/MK028/MK045/IT004/IT014/", UCase(UserID)) > 0 Then
            BAdmin.Visible = True
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        GridView1.Visible = False
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        '
        Dim SQL As String
        Dim xPuller As String
        Dim xShort As String

        GridView1.Visible = False
        GridView2.Visible = False
        DRDImage.Visible = False

        'Hide Register Field
        HideRegisterField()
        '--
        '-- IRW LINK
        pSize = Request.QueryString("pSize")                'Size
        pFamily = Request.QueryString("pFamily")            'family
        pSlider = Request.QueryString("pSlider")            'Slider
        pBuyer = Request.QueryString("pBuyer")              'Buyer
        pSource = Request.QueryString("pSource")             'Source
        DKSource.Text = ""
        DKPuller.Text = ""
        '
        'If pBuyer <> "" Then DKBuyer.Text = pBuyer
        If pSlider <> "" Then
            DKSource.Text = pSource
            '
            '外部 IRW SYSTEM連入
            If pSource = "IRW" Then
                xPuller = ""
                xShort = ""
                ' SHORT PULLER
                SQL = "select top 1 Puller, Short "
                SQL = SQL & "from M_PullerShortDetail "
                SQL = SQL & "where BUYER <> 'zzzzzz' "
                SQL = SQL & "AND '" & pSlider & "' like '%' + Short + '%' "
                SQL = SQL & "Order by len(short) desc "
                Dim dt_Short As DataTable = uDataBase.GetDataTable(SQL)
                If dt_Short.Rows.Count > 0 Then
                    xPuller = Trim(dt_Short.Rows(0).Item("Puller").ToString)
                    xShort = Trim(dt_Short.Rows(0).Item("Short").ToString)
                End If
                '
                'MsgBox(pSlider & ":" & xShort & ":" & xPuller)
                ' Get Puller + Color Inf.
                SQL = "select top 1 Puller, Puller+Color As Slider "
                SQL = SQL & "from V_PullerColorEDX "
                SQL = SQL & "where '" & Replace(pSlider, xShort, xPuller) & "' like '%'+ Puller+Color "
                SQL = SQL & "or '" & pSlider & "' like '%'+ Puller+Color "
                SQL = SQL & "Order by puller, len(puller+color) desc, color desc "
                '
                Dim dt_Slider As DataTable = uDataBase.GetDataTable(SQL)
                If dt_Slider.Rows.Count > 0 Then
                    DKPuller.Text = Trim(dt_Slider.Rows(0).Item("Puller").ToString)
                    DKOther.Text = Trim(dt_Slider.Rows(0).Item("Slider").ToString)
                Else
                    DKOther.Text = ""
                End If
            End If
            '
            '外部 ISOS SYSTEM連入
            If pSource = "ISOS" Then
                DKBuyer.Text = pBuyer
                DKPullerCode.Text = pSlider
            End If
            '
            '自已 ISIP SYSTEM連入
            If pSource = "ISIP" Then
                DKBuyer.Text = pBuyer
                DKOther.Text = pSlider
            End If
            '
            'ImagePortal
            If pSource = "IMGP" Then
                DKPullerCode.Text = pSlider
                DKColor.Text = pBuyer
            End If
        Else
            '
            '外部 ISOS SYSTEM連入
            If pSource = "ISOS" Then
                DKBuyer.Text = pBuyer
                DKPullerCode.Text = pSlider
            End If
        End If
        '
        AtSPColor.Checked = False
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

        DKBuyer.Text = UCase(DKBuyer.Text)

        SQL = "select top 150 "
        SQL = SQL & "[Puller], [Color], [ColorName], [BYColorCode], [Buyer], [BuyerName], [DTM_YOBI1], "
        SQL = SQL & "[OR_YOBI2] AS TapeColor, "
        '
        SQL = SQL & "CASE WHEN [DTM_YOBI1]='X' THEN "
        SQL = SQL & "  CASE WHEN [Remark]<>'' THEN LTRIM(RTRIM([DTM_YOBI2])) + '|' + [Remark] ELSE LTRIM(RTRIM([DTM_YOBI2])) END "
        SQL = SQL & "ELSE [Remark] "
        SQL = SQL & "END AS [Remark], "
        '
        SQL = SQL & "[DTM], [IRW], [ORDERS], "
        '
        SQL = SQL & "CASE WHEN RD_YOBI1>'1' THEN RD+' *' ELSE RD END AS RD, "
        SQL = SQL & "CASE WHEN RD_YOBI1>'1' THEN 'http://10.245.1.6/ISOSQC/ManySupplierList.aspx?pUserID=' + '" & UserID & "' + '&pPuller=' + Puller + '&pColor=' + Color + '&pFun=MANYRD' ELSE '' END AS RDMUrl, "
        SQL = SQL & "CASE WHEN EDX_YOBI1>'1' THEN EDX+' *' ELSE EDX END AS EDX, "
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
        SQL = SQL & "CASE WHEN ISNULL((SELECT TOP 1 [ImagePath] FROM [M_RDPullerImage] WHERE [SliderGRCode] =V_PullerColorEDX.PULLER + V_PullerColorEDX.COLOR AND FORMNO IN ('008002')), '') = '' THEN '' "
        SQL = SQL & "     ELSE 'IMG' "
        SQL = SQL & "END as EDX_IMG, "
        SQL = SQL & "'http://10.245.1.6/ISOSQC/Http2File.aspx?pUserID=' + '" & UserID & "' + '&pNo=' + [EDX_NO] as EDXIMGUrl, "
        '
        SQL = SQL & "Case When [IRW_NO]<>'' then [IRW_NO] else '' end as IRW_NO, "
        SQL = SQL & "'http://10.245.1.6/IRW/ItemRegisterSheet_02.aspx?pFormNo=001151&pFormSno=' + ltrim(rtrim(str(IRW_formsno))) as IRWUrl, "
        SQL = SQL & "[IRW_AppDate], "
        '
        SQL = SQL & "Case When [OR_NO]<>'' then [OR_NO] else '' end as OR_NO, "
        SQL = SQL & "'http://10.245.1.6/WorkFlowSub/PC_INQWingsOrder.aspx?pOrderNo=' + OR_NO + '&pPuller=' + Puller + Color as ORUrl, "
        SQL = SQL & "[OR_CfmDate], "

        SQL = SQL & "Case When [OR_NO]<>'' then 'IMG' else '' end as OR_IMG, "
        SQL = SQL + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & UserID & "&pBuyer=" & "" & "&pBuyerItem=' + LTrim(RTrim(OR_YOBI1)) + '&pFun=IMG" & "' as ORIMGURL, "
        SQL = SQL & "[PullerKey] "
        '
        SQL = SQL & "from V_PullerColorEDX "
        SQL = SQL & "where puller + color <> 'zzzzzzzz' "
        ' Puller
        If DKPullerCode.Text <> "" Then
            If InStr("/5N18/6N18/7N18/", "/" & pSlider & "/") > 0 Then
                SQL = SQL & "and puller LIKE '_N18' "
            Else
                If InStr("/N5N18/N6N18/N7N18/", "/" & pSlider & "/") > 0 Then
                    SQL = SQL & "and puller LIKE 'N_N18' "
                Else
                    SQL = SQL & "and puller = '" & DKPullerCode.Text & "' "
                End If
            End If
        End If
        ' Color
        If DKColor.Text <> "" Then
            SQL = SQL & "and Color = '" & DKColor.Text & "' "
        Else
            If AtSPColor.Checked = True Then
                SQL = SQL & "and Color = '' "
            End If
        End If
        ' Buyer
        If DKBuyer.Text <> "" Then
            If Len(DKBuyer.Text) = 6 Then
                SQL = SQL & "and buyer like '%" & DKBuyer.Text & "%' "
            Else
                SQL = SQL & "and buyer = '" & DKBuyer.Text & "' "
            End If
        End If
        '' Buyer
        'If DKBuyer.Text <> "" Then
        '    SQL = SQL & "and buyer + buyername like '%" & DKBuyer.Text & "%' "
        'End If
        ' Other
        If DKOther.Text <> "" Then
            'SQL = SQL & "and [Puller]+[Color]+[ColorName]+[BYColorCode]+[Buyer]+[BuyerName]+[DTM_YOBI1]+[DTM_YOBI2] like '%" & DKOther.Text & "%' "
            SQL = SQL & "and [Puller]+[Color]+[ColorName]+[BYColorCode]+[DTM_YOBI1]+[DTM_YOBI2]+[OR_YOBI2] like '%" & DKOther.Text & "%' "
        End If
        '
        SQL = SQL & "Order by puller, len(puller+color) desc, color desc "
        '
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
        '外部 IRW SYSTEM連入
        If DKSource.Text = "IRW" Then
            SQL = "SELECT "
            SQL &= "CASE WHEN MMS=1 THEN '模廢' ELSE '' END AS MMSDESC, "
            SQL &= "Puller, SliderCode, Spec, SUPPLIER, Sts AS Status, "

            SQL &= "CASE WHEN RDNO<>'' THEN RDNO ELSE '' END AS RDNOM, "
            SQL &= "CASE WHEN RDNO<>'' THEN 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FORMSNO))) ELSE '' END AS RDNOUrl, "

            SQL &= "CASE WHEN OPContact=1 THEN 'L' ELSE '' END AS OPContactM, "
            SQL &= "CASE WHEN OPContact=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FORMSNO))) + '&pFun=OP' ELSE '' END AS OPContactL, "

            SQL &= "CASE WHEN Contact=1 THEN 'L' ELSE '' END AS ContactM, "
            SQL &= "CASE WHEN Contact=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pFun=CT' ELSE '' END AS ContactL, "

            SQL &= "CASE WHEN SliderDetail=1 THEN 'L' ELSE '' END AS SliderDetailM, "
            SQL &= "CASE WHEN SliderDetail=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pCode=' + PULLER+COLOR + '&pFun=SD' ELSE '' END AS SliderDetailL, "

            SQL &= "CASE WHEN Surface=1 THEN 'L' ELSE '' END AS SurfaceM, "
            SQL &= "CASE WHEN Surface=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pFun=SF' ELSE '' END AS SurfaceL, "

            SQL &= "CASE WHEN ColorAppend=1 THEN 'L' ELSE '' END AS ColorAppendM, "
            SQL &= "CASE WHEN ColorAppend=1 THEN 'http://10.245.1.10/WorkFlow/ManufOutList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) + '&pFun=CA' ELSE '' END AS ColorAppendL, "

            SQL &= "CASE WHEN YKKGroupCopy=1 THEN 'L' ELSE '' END AS YKKGroupCopyM, "
            SQL &= "CASE WHEN YKKGroupCopy=1 THEN 'http://10.245.1.6//SPDSub/YKKGroupCopyList.aspx?pFormNo='+ FormNo + '&pFormSno=' + RTRIM(LTRIM(STR(FormSno))) ELSE '' END AS YKKGroupCopyL "
            '
            SQL = SQL & "FROM V_PullerColor2SPD "
            '
            If InStr("/5N18/6N18/7N18/", "/" & Trim(DKPuller.Text) & "/") > 0 Then
                SQL = SQL & "WHERE puller LIKE '_N18' "
            Else
                If InStr("/N5N18/N6N18/N7N18/", "/" & Trim(DKPuller.Text) & "/") > 0 Then
                    SQL = SQL & "WHERE puller LIKE 'N_N18' "
                Else
                    SQL = SQL & "where puller = '" & Trim(DKPuller.Text) & "' "
                End If
            End If
            '
            GridView2.Visible = True
            GridView2.DataSource = uDataBase.GetDataTable(SQL)
            GridView2.DataBind()

            AtCloseRDW.Style("left") = 1120 & "px"
            AtCloseRDW.Checked = False
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Row Command
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim idx As Integer
        Dim str, Cmd, sql As String
        Dim xPullerKey, xPuller, xColor As String
        'Puller+Color
        idx = Convert.ToInt32(e.CommandArgument)
        xPullerKey = GridView1.DataKeys(idx).Value.ToString()
        'Puller
        str = xPullerKey.Trim
        xPuller = Mid(str, 1, InStr(str, "/") - 1)
        'Color
        str = Mid(str, InStr(str, "/") + 1, 99)
        xColor = Mid(str, 1, InStr(str, "/") - 1)
        'MsgBox(xPullerKey & "-" & xPuller & "-" & xColor)
        '
        'EDX IMAGE JPG & PDF
        If e.CommandName = "EDXIMG" Then
            str = "\\10.245.1.6\www$\ISOSQC\Document\EDX\" & xPuller.Trim & xColor.Trim & ".jpg"
            str = Replace(UCase(str), "N*55", "Nx55")
            str = Replace(UCase(str), "N*63", "Nx63")
            str = Replace(UCase(str), "N*73", "Nx73")
            If File.Exists(str) Then
                DEDXImage.ImageUrl = str
                DEDXImage.Visible = True
                '
                AtCloseEDXW.Style("left") = 1120 & "px"
                AtCloseEDXW.Checked = False
                '
                HideRegisterField()
                GridView2.Visible = False
            Else
                sql = "select EDX_NO "
                sql = sql & "from V_PullerColorEDX "
                sql = sql & "where PullerKey = '" & xPullerKey.Trim & "' "
                sql = sql & "  and EDX_NO <> 'QC5Y' "
                sql = sql & "Order by EDX_NO "
                Dim dt_EDXNO As DataTable = uDataBase.GetDataTable(sql)
                If dt_EDXNO.Rows.Count > 0 Then
                    Cmd = "<script>" + _
                                "window.open('http://10.245.1.6/ISOSQC/Http2File.aspx?pUserID=" & UserID & "&pNo=" & dt_EDXNO.Rows(0).Item("EDX_NO").trim & "','',''); " + _
                          "</script>"
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     New Register
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim SQL As String
        '
        'Puller Code
        DPullerCode.Text = Replace(GridView1.SelectedRow.Cells(2).Text, "&nbsp;", "")
        '
        'ColorName
        SQL = "select top 1 NextColorName "
        SQL = SQL & "from V_PullerNextColor "
        If DPullerCode.Text = "DRR1" Or DPullerCode.Text = "DRR2" Then
            SQL = SQL & "where puller IN ('DRR1','DRR2') "
        Else
            SQL = SQL & "where puller = '" & DPullerCode.Text & "' "
        End If
        SQL = SQL & "Order by [NextColorName] DESC "
        '
        Dim dt_NextColor As DataTable = uDataBase.GetDataTable(SQL)
        If dt_NextColor.Rows.Count > 0 Then
            DColor.Text = dt_NextColor.Rows(0).Item("NextColorName").ToString
            '
            If DColor.Text = "A" Or DColor.Text = "B" Or DColor.Text = "C" Then
                DColor.Text = "D"
            End If
        Else
            DColor.Text = ""
        End If
        'ColorName
        DColorName.Text = Replace(GridView1.SelectedRow.Cells(4).Text, "&nbsp;", "")
        If DColorName.Text = "" Then DColorName.Text = "--"
        '23/12/29 COLORNAME & BUYER COLOR CODE DEFAULT = BLANK
        DColorName.Text = ""

        'Buyer Color Code
        DBYColorCode.Text = Replace(GridView1.SelectedRow.Cells(5).Text, "&nbsp;", "")
        If DBYColorCode.Text = "" Then DBYColorCode.Text = "--"
        '23/12/29 COLORNAME & BUYER COLOR CODE DEFAULT = BLANK
        DBYColorCode.Text = ""

        'Buyer
        DBuyer.Text = Replace(GridView1.SelectedRow.Cells(6).Text, "&nbsp;", "")
        'BuyerName
        DBuyerName.Text = Replace(GridView1.SelectedRow.Cells(7).Text, "&nbsp;", "")

        'TapeColor  DEFAULT = BLANK
        DTapeColor.Text = ""
        'Remark
        DRemark.Text = Replace(GridView1.SelectedRow.Cells(9).Text, "&nbsp;", "")
        '23/12/29 COLORNAME & BUYER COLOR CODE DEFAULT = BLANK
        DRemark.Text = ""
        '
        'Show Register Field
        ShowRegisterField()
        '
        AtBColor.Checked = False
        AtCColor.Checked = False
        'Hide Gridview2
        GridView2.Visible = False
        DRDImage.Visible = False
        '
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
        If InStr("/5N18/6N18/7N18/", "/" & Replace(GridView1.Rows(e.NewEditIndex).Cells(2).Text, "&nbsp;", "") & "/") > 0 Then
            SQL = SQL & "WHERE puller LIKE '_N18' "
        Else
            If InStr("/N5N18/N6N18/N7N18/", "/" & Replace(GridView1.Rows(e.NewEditIndex).Cells(2).Text, "&nbsp;", "") & "/") > 0 Then
                SQL = SQL & "WHERE puller LIKE 'N_N18' "
            Else
                SQL = SQL & "where puller = '" & Replace(GridView1.Rows(e.NewEditIndex).Cells(2).Text, "&nbsp;", "") & "' "
            End If
        End If
        '
        GridView2.Visible = True
        GridView2.DataSource = uDataBase.GetDataTable(Sql)
        GridView2.DataBind()
        '
        AtCloseRDW.Style("left") = 1120 & "px"
        AtCloseRDW.Checked = False
        '
        HideRegisterField()
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
        Dim lnk As HyperLink = GridView1.Rows(e.RowIndex).Cells(16).Controls(0)
        '
        SQL = "select top 1 [ImagePath] as Path, [Yobi1] as BackupPath "
        SQL = SQL & "from M_RDPullerImage "
        SQL = SQL & "where No = '" & Replace(lnk.Text, "&nbsp;", "") & "' "
        SQL = SQL & "Order by createtime "
        Dim dt_RDImage As DataTable = uDataBase.GetDataTable(SQL)
        If dt_RDImage.Rows.Count > 0 Then
            '
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
            AtCloseRDW.Style("left") = -500 & "px"
            AtCloseRDW.Checked = False
            '
            HideRegisterField()
            GridView2.Visible = False
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Register Field Control 不顯示
    '**
    '*****************************************************************
    Sub HideRegisterField()
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        textbox5.Visible = False
        textbox6.Visible = False
        textbox7.Visible = False
        textbox9.Visible = False
        textbox10.Visible = False
        textbox13.Visible = False
        '
        DPullerCode.Visible = False
        DColor.Visible = False
        DColorName.Visible = False
        DBYColorCode.Visible = False
        DBuyer.Visible = False
        DBuyerName.Visible = False
        DTapeColor.Visible = False
        DRemark.Visible = False
        AtBColor.Visible = False
        AtCColor.Visible = False
        '
        BRegister.Visible = False
        '
        GridView2.Visible = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Register Field Control 顯示 +  PULLER, COLOR權限判定
    '**
    '*****************************************************************
    Sub ShowRegisterField()
        TextBox2.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        textbox5.Visible = True
        textbox6.Visible = True
        textbox7.Visible = True
        textbox9.Visible = True
        textbox10.Visible = True
        textbox13.Visible = True
        '
        DPullerCode.Visible = True
        DColor.Visible = True
        DColorName.Visible = True
        DBYColorCode.Visible = True
        DBuyer.Visible = True
        DBuyerName.Visible = True
        DTapeColor.Visible = True
        DRemark.Visible = True
        AtBColor.Visible = True
        AtCColor.Visible = True
        '
        BRegister.Visible = True
        '
        GridView2.Visible = False
        '
        ' PULLER, COLOR權限
        '
        ' COMMON USER:PULLER + COLOR=不可修改
        DPullerCode.ReadOnly = True
        DColor.ReadOnly = True
        ' SUPER USER:PULLER = COLOR可修改
        Dim SQL As String
        SQL = "select top 1 UserName "
        SQL = SQL & "from V_NewPullerColorUser "
        SQL = SQL & "where Cat IN ('SUPER') "
        SQL = SQL & "AND UserID = '" & UCase(UserID) & "' "
        SQL = SQL & "Order by UserName "
        Dim dt_UserNewColor As DataTable = uDataBase.GetDataTable(SQL)
        If dt_UserNewColor.Rows.Count > 0 Then
            DPullerCode.ReadOnly = False
            DColor.ReadOnly = False
        Else
            '
            ' 特定PULLER可修改 COLOR
            ' 發行非流水號COLOR COLOR可修改
            SQL = "select top 1 UserName "
            SQL = SQL & "from V_NewPullerColorUser "
            SQL = SQL & "where Cat IN ('PULLER') "
            SQL = SQL & "AND UserID = '" & UCase(UserID) & "' "
            SQL = SQL & "AND BuyerPullerStr Like '%" & Trim(DPullerCode.Text) & "/%' "
            SQL = SQL & "Order by UserName "
            Dim dt_UserNewPuller As DataTable = uDataBase.GetDataTable(SQL)
            If dt_UserNewPuller.Rows.Count > 0 Then
                DPullerCode.ReadOnly = True
                DColor.ReadOnly = False
            End If
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Register Field 清初值 
    '**
    '*****************************************************************
    Sub BlankRegisterField()
        DPullerCode.Text = ""
        DColor.Text = ""
        DColorName.Text = ""
        DBYColorCode.Text = ""
        DBuyer.Text = ""
        DBuyerName.Text = ""
        DTapeColor.Text = ""
        DRemark.Text = ""
        AtBColor.Checked = False
        AtCColor.Checked = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Append Puller Color Master
    '**
    '*****************************************************************
    Protected Sub BRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRegister.Click
        Dim sql As String = ""
        Dim xNewColor As Boolean
        Dim xSharePuller As Boolean
        '
        xNewColor = False
        xSharePuller = False
        '
        '共用PULLER
        sql = "select top 1 Puller "
        sql = sql & "from M_ExcludePullerColor "
        sql = sql & "where Puller = '" & Trim(DPullerCode.Text) & "' "
        Dim dt_SharePuller As DataTable = uDataBase.GetDataTable(sql)
        If dt_SharePuller.Rows.Count > 0 Then
            xSharePuller = True
        End If
        '
        'CHECK INPUT DATA
        If DPullerCode.Text = "" Or DColor.Text = "" Or DColorName.Text = "" Or DBYColorCode.Text = "" Or DBuyer.Text = "" Or DBuyerName.Text = "" Then
            xNewColor = True
            uJavaScript.PopMsg(Me, "  [Tape Color], [Remak]以外都必須輸入 !! ")
        Else
            '
            'PULLER = YG/TW239/Y576/Y616/Y250/Y582  TAPE COLOR=須輸入
            If InStr("YG/TW239/Y576/Y616/Y250/Y582/", DPullerCode.Text & "/") > 0 Then
                '
                If DTapeColor.Text = "" Then
                    xNewColor = True
                    uJavaScript.PopMsg(Me, "  Puller " & DPullerCode.Text & " 時 [Tape Color] 必須輸入 !! ")
                End If
            End If
        End If
        '
        'CHECK USER AUTHORITY
        If xNewColor = False Then
            sql = "select top 1 UserName "
            sql = sql & "from V_NewPullerColorUser "
            sql = sql & "where Cat IN ('BUYER','SUPER') "
            sql = sql & "AND UserID = '" & UCase(UserID) & "' "
            sql = sql & "AND ( "
            sql = sql & "  BuyerPullerStr Like '%" & Trim(DBuyer.Text) & "/%' "
            sql = sql & "  OR "
            sql = sql & "  BuyerPullerStr Like '%" & "ALL" & "/%' "
            sql = sql & " ) "
            sql = sql & "Order by UserName "
            Dim dt_UserNewColor As DataTable = uDataBase.GetDataTable(sql)
            If dt_UserNewColor.Rows.Count <= 0 Then
                xNewColor = True
                uJavaScript.PopMsg(Me, "  偵測到您非此BUYER擔當者，請確認Buyer !! ")
            End If
        End If
        '
        '移行未補齊
        If xNewColor = False Then
            sql = "select Top 1 Puller "
            sql = sql & "from M_PullerColor "
            sql = sql & "where Puller = '" & Trim(DPullerCode.Text) & "' "
            sql = sql & "AND   Buyer  = '" & Trim(DBuyer.Text) & "' "
            Dim dt_PullerColor As DataTable = uDataBase.GetDataTable(sql)
            If dt_PullerColor.Rows.Count <= 0 Then
                '
                sql = "select Top 1 Slider1, Slider2, BuyerCode, No "
                sql = sql & "from F_ItemRegisterSheet "
                sql = sql & "where STS IN (0,1) "
                sql = sql & "and ( "
                sql = sql & "      Slider1 like '%" & Trim(DPullerCode.Text) & "%' "
                sql = sql & "   or Slider2 like '%" & Trim(DPullerCode.Text) & "%' "
                sql = sql & " ) "
                sql = sql & "AND   BuyerCode  = '" & Trim(DBuyer.Text) & "' "
                sql = sql & "Order by Date desc "
                Dim dt_IRWSheet As DataTable = uDataBase.GetDataTable(sql)
                If dt_IRWSheet.Rows.Count > 0 Then
                    xNewColor = True
                    uJavaScript.PopMsg(Me, dt_IRWSheet.Rows(0).Item("No") & ":" & dt_IRWSheet.Rows(0).Item("BuyerCode") & "-" & dt_IRWSheet.Rows(0).Item("Slider1") & "/" & dt_IRWSheet.Rows(0).Item("Slider2") & "  偵測到未補齊移行舊資料，請確認 !! ")
                End If
            End If
        End If
        '
        '存在(ISIP)或未開放
        If xNewColor = False Then
            sql = "select [Puller], [Color], [ColorName] "
            sql = sql & "from V_PullerColorList "
            sql = sql & "WHERE puller + color = '" & Trim(DPullerCode.Text) & Trim(DColor.Text) & "' "
            sql = sql & " OR  puller + color = '" & Trim(DPullerCode.Text) & Trim("ZZZZ") & "' "
            Dim dtPuller As DataTable = uDataBase.GetDataTable(sql)
            If dtPuller.Rows.Count > 0 Then
                xNewColor = True
                uJavaScript.PopMsg(Me, Trim(DPullerCode.Text) & "-" & Trim(DColor.Text) & "  已存在或未開放 !! ")
            End If
            '
            If xNewColor = False Then
                sql = "select [Puller], [Color], [ColorName] "
                sql = sql & "from V_PullerColorList "
                sql = sql & "WHERE puller = '" & Trim(DPullerCode.Text) & "' "
                sql = sql & " AND REPLACE(ColorName,' ','') = '" & Replace(Trim(DColorName.Text), " ", "") & "' "
                Dim dtPullerA As DataTable = uDataBase.GetDataTable(sql)
                If dtPullerA.Rows.Count > 0 Then
                    xNewColor = True
                    uJavaScript.PopMsg(Me, Trim(DColorName.Text) & "  已存在 !! ")
                End If
            End If
        End If
        '
        '存在(EDX)
        If xNewColor = False Then
            sql = "select [Puller], [Color] "
            sql = sql & "from V_EDXQA2ISIP "
            sql = sql & "WHERE puller + color = '" & Trim(DPullerCode.Text) & Trim(DColor.Text) & "' "
            Dim dtEDX As DataTable = uDataBase.GetDataTable(sql)
            If dtEDX.Rows.Count > 0 Then
                xNewColor = True
                uJavaScript.PopMsg(Me, Trim(DPullerCode.Text) & "-" & Trim(DColor.Text) & "  已存在(EDX) !! ")
            End If
        End If
        '
        'PULLER CONTROL 
        If xNewColor = False Then
            sql = "select [ColorName] "
            sql = sql & "from M_PullerColorControl "
            sql = sql & "WHERE [ColorName] = '" & Trim(DColor.Text) & "' "
            'YG / YN82R / RHN / KLPR
            If Trim(DPullerCode.Text) = "YG" Or Trim(DPullerCode.Text) = "YN82R" Or Trim(DPullerCode.Text) = "RHN" Or Trim(DPullerCode.Text) = "KLPR" Then
                sql = sql & "and Cat = '" & Trim(DPullerCode.Text) & "' "
            Else
                sql = sql & "and Cat = 'ALL' "
            End If
            '
            Dim dtControl As DataTable = uDataBase.GetDataTable(sql)
            If dtControl.Rows.Count <= 0 Then
                xNewColor = True
                uJavaScript.PopMsg(Me, Trim(DPullerCode.Text) & "-" & Trim(DColor.Text) & "  非控制中心無法發行新色號 !! ")
            End If
        End If
        '
        'REGISTER PULL DATA
        If xNewColor = False Then
            '
            sql = "INSERT INTO M_PullerColor ( "
            sql = sql & "[Puller],[Color],[ColorName],[BYColorCode],[Buyer],[BuyerName],[Remark], "
            sql = sql & "[RD], [DTM], [EDX], [IRW], [ORDERS], "
            sql = sql & "[DTM_YOBI1], [DTM_YOBI2], "
            sql = sql & "[RD_NO],[RD_AppDate],[RD_Supplier],[RD_FormNo],[RD_FormSno],[RD_YOBI1],[RD_YOBI2], "
            sql = sql & "[EDX_NO],[EDX_AppDate],[EDX_Supplier],[EDX_FormNo],[EDX_FormSno],[EDX_YOBI1],[EDX_YOBI2], "
            sql = sql & "[IRW_NO],[IRW_AppDate],[IRW_FormNo],[IRW_FormSno],[IRW_YOBI1],[IRW_YOBI2], "
            sql = sql & "[OR_NO],[OR_CfmDate],[OR_Customer],[OR_YOBI1],[OR_YOBI2], "
            sql = sql & "[Yobi1],[Yobi2],[Yobi3],[Yobi4], "
            sql = sql & "[CreateUser],[CreateTime],[ModifyUser], [ModifyTime] "
            sql = sql & ") "
            sql = sql & "VALUES( "
            ' BUYER PULLER
            sql &= "'" & DPullerCode.Text & "' ,"
            sql &= "'" & DColor.Text & "' ,"
            sql &= "'" & DColorName.Text & "' ,"
            sql &= "'" & DBYColorCode.Text & "' ,"
            sql &= "'" & DBuyer.Text & "' ,"
            sql &= "'" & UCase(DBuyerName.Text) & "' ,"
            sql &= "'" & DRemark.Text & "' ,"
            ' STATUS
            sql &= "'" & "WIP" & "' ,"
            sql &= "'" & "WIP" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            ' DTM
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            ' RD
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= " " & "0" & " ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            ' EDX
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= " " & "0" & " ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            ' IRW
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= " " & "0" & " ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            ' OR
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & DTapeColor.Text & "' ,"
            ' YOBI
            sql &= "'" & "ISIP" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            sql &= "'" & "" & "' ,"
            '
            sql &= "'" & Request.QueryString("pUserID") & "' ,"
            sql &= "'" & NowDateTime & "' ,"
            sql &= "'" & Request.QueryString("pUserID") & "' ,"
            sql &= "'" & NowDateTime & "' )"
            '
            uDataBase.ExecuteNonQuery(sql)
            '
            'Blank & Hide Filed 
            BlankRegisterField()
            HideRegisterField()
            'Rrefresh
            ShowData()
            '
            uJavaScript.PopMsg(Me, "  已註冊成功 !! ")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
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
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "" & "<BR>" & ""
            tcl(0).BackColor = Color.White
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "" & "<BR>" & ""
            tcl(1).BackColor = Color.White
            'BUYER PULLER
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Puller" & "<BR>" & "Code"
            tcl(2).BackColor = Color.Blue
            'Puller Code

            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Color" & "<BR>" & "Code"
            tcl(3).BackColor = Color.Blue
            'Color Code	

            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Buyer Color" & "<BR>" & "Name"
            tcl(4).BackColor = Color.Blue
            'Buyer Color Name

            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Buyer Color" & "<BR>" & "Code"
            tcl(5).BackColor = Color.Blue
            'Buyer Color Code

            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Buyer" & "<BR>" & ""
            tcl(6).BackColor = Color.Blue
            'Buyer

            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Buyer" & "<BR>" & "Name"
            tcl(7).BackColor = Color.Blue
            'Buyer Name

            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "Tape " & "<BR>" & "Color"
            tcl(8).BackColor = Color.Blue
            'Tape Color

            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "Remark" & "<BR>" & ""
            tcl(9).BackColor = Color.Blue
            'Remark

            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "Drop" & "<BR>" & ""
            tcl(10).BackColor = Color.Blue
            '
            'STATUS
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "RD" & "<BR>" & ""
            tcl(11).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            'tcl(12).Text = "DTM" & "<BR>" & ""
            tcl(12).Text = "DTM" & "<BR>" & "dev. date"
            tcl(12).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "EDX" & "<BR>" & ""
            tcl(13).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "IRW" & "<BR>" & ""
            tcl(14).BackColor = Color.Purple
            tcl.Add(New TableHeaderCell())
            'tcl(15).Text = "ORDERS" & "<BR>" & ""
            tcl(15).Text = "OR" & "<BR>" & ""
            tcl(15).BackColor = Color.Purple
            'LINK
            tcl.Add(New TableHeaderCell())
            tcl(16).Text = "RD" & "<BR>" & "NO."
            tcl(16).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(17).Text = "RD" & "<BR>" & "App.d"
            tcl(17).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(18).Text = "RD" & "<BR>" & "Supplier"
            tcl(18).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(19).Text = "RD" & "<BR>" & "Images"
            tcl(19).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(20).Text = "EDX" & "<BR>" & "NO."
            tcl(20).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(21).Text = "EDX" & "<BR>" & "App.d"
            tcl(21).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(22).Text = "EDX" & "<BR>" & "Supplier"
            tcl(22).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(23).Text = "EDX" & "<BR>" & "Images"
            tcl(23).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(24).Text = "IRW" & "<BR>" & "NO."
            tcl(24).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(25).Text = "IRW" & "<BR>" & "App.d"
            tcl(25).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(26).Text = "OR" & "<BR>" & "NO."
            tcl(26).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(27).Text = "OR" & "<BR>" & "Sales.d"
            tcl(27).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(28).Text = "OR" & "<BR>" & "Images"
            tcl(28).BackColor = Color.Green
            '
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "Puller Color"
            H3tc1.ColumnSpan = 11
            H3row.Cells.Add(H3tc1)
            '
            Dim H4tc1 As TableCell = New TableCell
            H4tc1.Text = "Progress Status"
            H4tc1.ColumnSpan = 5
            H3row.Cells.Add(H4tc1)
            '
            Dim H5tc1 As TableCell = New TableCell
            H5tc1.Text = "Link"
            H5tc1.ColumnSpan = 14
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
        Dim i, j As Integer
        Dim str, str1 As String
        Dim BuyerInf, TColorInf As String()
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(2).ForeColor = Color.Red
            e.Row.Cells(3).ForeColor = Color.Red
            '
            e.Row.Cells(4).Font.Bold = True
            e.Row.Cells(5).Font.Bold = True
            '
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
                    ' 8:tape color
                    If e.Row.Cells(8).Text <> "&nbsp;" Then
                        BuyerInf = e.Row.Cells(8).Text.Split("|")
                        e.Row.Cells(8).Text = BuyerInf(j)
                    End If
                    ' 9:remark
                    If InStr(e.Row.Cells(6).Text, "|") > 0 Then
                        BuyerInf = e.Row.Cells(9).Text.Split("|")
                        e.Row.Cells(9).Text = BuyerInf(j)
                    End If
                Else
                    If Replace(e.Row.Cells(8).Text, "|", "") <> "" And e.Row.Cells(8).Text <> "&nbsp;" Then
                        ' 6:buyer
                        'str = ""
                        'BuyerInf = e.Row.Cells(9).Text.Split("|")
                        'str1 = e.Row.Cells(8).Text & "||||"
                        'TColorInf = str1.Split("|")
                        '
                        'For j = 0 To BuyerInf.Length - 1
                        '    If TColorInf(j) <> "&nbsp;" And TColorInf(j) <> "" Then
                        '        If InStr(e.Row.Cells(6).Text, "|") > 0 Then
                        '            e.Row.Cells(9).Text = str & "|" & BuyerInf(j)
                        '        Else
                        '            e.Row.Cells(9).Text = str & "|" & BuyerInf(j) & "|"
                        '        End If
                        '        str = e.Row.Cells(9).Text
                        '    Else
                        '        e.Row.Cells(9).Text = str & "|" & BuyerInf(j)
                        '        str = e.Row.Cells(9).Text
                        '    End If
                        'Next
                        '
                    End If
                    '
                End If
            Next
            '
            ' 多BUYER換行
            For i = 4 To 8
                e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, "|", "<br>")
            Next
            '
            ' REMARK調正
            'If Mid(e.Row.Cells(9).Text, 1, 1) = "|" Then
            '    e.Row.Cells(9).Text = Mid(e.Row.Cells(9).Text, 2, 99)
            'End If
            If Len(e.Row.Cells(6).Text) > 6 Then
                e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, "|", "<br>")
            Else
                e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, "|", "")
            End If
            ' Drop Status
            For i = 10 To 10
                If e.Row.Cells(i).Text = "X" Then
                    e.Row.Cells(i).BackColor = Color.Red
                    e.Row.Cells(i).ForeColor = Color.White
                    e.Row.Cells(i).Font.Bold = True
                End If
            Next
            ' Progress Status
            For i = 11 To 17
                If InStr(e.Row.Cells(i).Text, "WIP") > 0 Then
                    e.Row.Cells(i).BackColor = Color.Cyan
                    e.Row.Cells(i).ForeColor = Color.Black
                    e.Row.Cells(i).Font.Bold = True
                End If
                '
                If i = 11 Then
                    Dim lnk As HyperLink = e.Row.Cells(i).Controls(0)
                    If InStr(lnk.Text, "WIP") > 0 Then
                        e.Row.Cells(i).BackColor = Color.Cyan
                        e.Row.Cells(i).ForeColor = Color.Black
                        e.Row.Cells(i).Font.Bold = True
                    End If
                End If
                '
                If i = 13 Then
                    Dim lnk As HyperLink = e.Row.Cells(i).Controls(0)
                    If InStr(lnk.Text, "WIP") > 0 Then
                        '從IRW進入ISIP, EDX如果是WIP,show品質分析依賴書page供參考(第一筆資料)
                        If pSource = "IRW" And e.Row.RowIndex = 0 Then
                            getQC()
                        End If
                        e.Row.Cells(i).BackColor = Color.Cyan
                        e.Row.Cells(i).ForeColor = Color.Black
                        e.Row.Cells(i).Font.Bold = True
                    End If
                End If
            Next
            ' RD Images lINK
            For i = 19 To 19
                If e.Row.Cells(i - 1).Text = "&nbsp;" Then e.Row.Cells(i).Text = ""
                e.Row.Cells(i).ForeColor = Color.Blue
            Next
            ' EDX Images lINK
            For i = 23 To 23
                str = CType(e.Row.Cells(i - 3).Controls(0), HyperLink).Text
                If str = "QC5Y" Then
                    e.Row.Cells(i).Text = ""
                Else
                    If e.Row.Cells(i - 1).Text = "&nbsp;" Then e.Row.Cells(i).Text = ""
                End If
                '
                e.Row.Cells(i).ForeColor = Color.Blue
            Next
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     B, C COLOR CHECKBOX
    '**
    '*****************************************************************
    Protected Sub AtBColor_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtBColor.CheckedChanged
        DColor.Text = "B"
        AtCColor.Checked = False
    End Sub

    Protected Sub AtCColor_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtCColor.CheckedChanged
        DColor.Text = "C"
        AtBColor.Checked = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO Puller Maint.)
    '**     
    '**
    '*****************************************************************
    Protected Sub BO365_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BO365.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('https://forms.office.com/r/4uvpdHxgaH');" + _
              "</script>"
        Response.Write(Cmd)
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
    '**(Close IMG Windows)
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
    '**(Advenced Images)
    '**     
    '**
    '*****************************************************************
    Protected Sub BADVImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADVImage.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NewAdvencedImages.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Admin Report)
    '**     
    '**
    '*****************************************************************
    Protected Sub BAdmin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BAdmin.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/AdminMenu.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '** 從IRW進入ISIP, EDX如果是WIP,show品質分析依賴書page供參考
    '**     
    '**
    '*****************************************************************
    Protected Sub getQC()
        Dim qcUrl As String
        'local test
        'qcUrl = "http://localhost:49231/ISOSQC/QCListinqCommission.aspx?pIRW=0&pSize=&pSlider=" & DKOther.Text & "&pPuller=&pWIP=1"
        qcUrl = "http://10.245.1.6/ISOSQC/QCListinqCommission.aspx?pIRW=0&pSize=&pSlider=" & DKOther.Text & "&pPuller=&pWIP=1"
        Response.Write("<script>window.open('" & qcUrl & "','');</script>")
    End Sub
End Class
