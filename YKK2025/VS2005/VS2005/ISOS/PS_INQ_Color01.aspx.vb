Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_Color01
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
    Dim oWaves As New Waves.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pBuyer As String            'Buyer
    Dim UserID As String            'UserID
    Dim pBColor As String           'BuyerColor
    Dim pFun As String              'SPC, SPP
    Dim pDTMWColor As String        'DTMWColor
    '
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")
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
        Response.Cookies("PGM").Value = "PS_INQ_Color01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        pBColor = Request.QueryString("pBColor")            'BuyerColor
        pFun = Request.QueryString("pFun")                  'SPC, SPP
        pDTMWColor = Request.QueryString("pDTMWColor")      'DTMColor
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        Label4.Style("left") = -500 & "px"
        LTeethPage.Style("left") = -500 & "px"
        LTapeDesc.Style("left") = -500 & "px"
        '
        DINQBUYERCOLORFile.Style("left") = -500 & "px"
        DINQBUYERCOLORFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQBUYERCOLOR_" & UserID & ".xlsm"
        '
        If pBuyer = "TW1741" Then
            Label4.Style("left") = 230 & "px"
        End If
        '
        '動作按鈕設定
        BSPInq.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('INQBUYERCOLORExcel')}"
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '
        If pBColor <> "" Then
            AtBuyer.Checked = True
            '
            AtFAS.Checked = False
            AtReplace.Checked = False
            AtSales.Checked = False
            AtDevlop.Checked = False
            AtFAS.Checked = False
        Else
            If pDTMWColor <> "" Then
                AtDevlop.Checked = True
                '
                AtBuyer.Checked = False
                AtFAS.Checked = False
                AtReplace.Checked = False
                AtSales.Checked = False
                AtFAS.Checked = False
            Else
                If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtDevlop.Checked = False Then
                    'AtFAS.Checked = True
                    AtBuyer.Checked = True
                    If pBuyer = "TW1741" Then
                        AtFAS.Checked = False
                        AtBuyer.Checked = True
                    End If
                    If pBuyer = "000098" Or pBuyer = "000068" Or pBuyer = "000151" Or pBuyer = "TW5068" Or pBuyer = "TW0655" Or pBuyer = "000141" Or pBuyer = "000053" Or pBuyer = "000055" Then
                        AtFAS.Checked = False
                        AtBuyer.Checked = True
                        AtFAS.Style("left") = -500 & "px"
                        AtReplace.Style("left") = -500 & "px"
                    End If
                    If pBuyer = "000141" Then
                        BSPInq.Style("left") = -500 & "px"
                    End If
                End If
            End If
        End If
        '
        DBUYER.Text = ""
        HBuyerCode.Text = ""
        DKEY1.Text = ""
        DKEY2.Text = ""
        '
        If pBuyer <> "" Then
            If pBuyer = "000021" Then DBUYER.Text = "TNF"
            If pBuyer = "000001" Then DBUYER.Text = "ADIDAS"
            If pBuyer = "000013" Then DBUYER.Text = "NIKE"
            If pBuyer = "000003" Then DBUYER.Text = "COLUMBIA"
            If pBuyer = "TW1741" Then DBUYER.Text = "LULULEMON"
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            If pBuyer = "000098" Then DBUYER.Text = "HELLY HANSEN"
            If pBuyer = "000068" Then DBUYER.Text = "BURTON"
            If pBuyer = "000151" Then DBUYER.Text = "PUMA"
            If pBuyer = "TW5068" Then DBUYER.Text = "HERSCHEL"
            If pBuyer = "TW0655" Then DBUYER.Text = "VERA BRADLEY"
            If pBuyer = "000141" Then DBUYER.Text = "PATAGONIA"
            If pBuyer = "000053" Then DBUYER.Text = "ARCTERYX"
            If pBuyer = "000055" Then DBUYER.Text = "SALOMON"

            If pBuyer = "TW1741" Or pBuyer = "000151" Then
                HBuyerCode.Text = "ISOS-" & pBuyer
            Else
                HBuyerCode.Text = "FALL-" & pBuyer
            End If

            If pBuyer = "TW5068" Or pBuyer = "TW0655" Or pBuyer = "000141" Then DColorType.Style("left") = -500 & "px"
        End If
        '
        DColorType.Items.Clear()
        'BC COLOR
        'TAPE
        Dim ListItem1 As New ListItem
        ListItem1.Text = "TAPE"
        ListItem1.Value = "TAPE"
        ListItem1.Selected = True
        DColorType.Items.Add(ListItem1)
        'PULLER
        Dim ListItem2 As New ListItem
        ListItem2.Text = "PULLER"
        ListItem2.Value = "PULLER"
        DColorType.Items.Add(ListItem2)
        '(限000013) TEETH & SLDFINISH 
        If pBuyer = "000013" Then
            Dim ListItem11 As New ListItem
            ListItem11.Text = "TEETH"
            ListItem11.Value = "TEETH"
            DColorType.Items.Add(ListItem11)
            '
            Dim ListItem12 As New ListItem
            ListItem12.Text = "SLDFINISH"
            ListItem12.Value = "SLDFINISH"
            DColorType.Items.Add(ListItem12)
        End If
        '
        'BCP(外部連動PULLER)
        If pFun = "BCP" Then
            If pBColor <> "" Then
                DKEY1.Text = pBColor
                DColorType.SelectedIndex = 1
            End If
        End If
        '
        'BCT(外部連動TAPE)
        If pFun = "BCT" Then
            If pBColor <> "" Then
                DKEY1.Text = pBColor
                DColorType.SelectedIndex = 0
                '
                '(限000013) 轉成TEETH
                If pBuyer = "000013" Then DColorType.SelectedIndex = 2
            End If
        End If
        '
        'DTMC COLOR
        If pDTMWColor <> "" Then DKEY1.Text = pDTMWColor
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        If pBuyer = "000013" And DColorType.SelectedValue = "TEETH" And AtBuyer.Checked = True Then
            LTeethPage.Style("left") = 630 & "px"
        Else
            LTeethPage.Style("left") = -500 & "px"
        End If

        If DColorType.SelectedValue = "PULLER" Then
            Dim Cmd, xSlider As String
            '
            xSlider = ""
            If DKEY1.Text <> "" Then xSlider = UCase(DKEY1.Text)
            If DKEY2.Text <> "" Then xSlider = xSlider & UCase(DKEY2.Text)
            '
            Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pSlider=" & xSlider & "&pSource=ISOS" & "','ISIP','status=1,toolbar=0,width=5000,height=3000,resizable=yes,scrollbars=yes,top=50, left=50');" + _
              "</script>"

            Response.Write(Cmd)
        Else
            ShowData()
        End If
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        'On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        If pBuyer = "TW1741" And DColorType.SelectedValue = "TAPE" And AtBuyer.Checked = True Then
            LTapeDesc.Style("left") = 40 & "px"
        Else
            LTapeDesc.Style("left") = -500 & "px"
        End If
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '=========================================================================
        '共用
        '=========================================================================
        '
        '--開發
        If AtDevlop.Checked = True Then
            sql = "SELECT top 300 "
            sql = sql + "'DTMW' as DTMW, "
            sql = sql + "'' as DTMWURL, "
            sql = sql + "'SLDC' as SLDC, "
            sql = sql + "'' as SLDCURL, "
            sql = sql + "'IRW' as IRW, "
            sql = sql + "'' as IRWURL, "
            '
            sql &= "YKKColorCode AS A1, "
            sql &= "Case When CustomerColorCode='' Then CustomerColor Else CustomerColor + '(' + CustomerColorCode + ')' End AS B1, "
            sql &= "Case When CustomerCode='' Then Customer Else Customer + '(' + CustomerCode +')' End AS C1, "
            sql &= "TapeType AS D1, "
            sql &= "No AS E1, Name AS F1, Sts AS G1, ApplyDate AS H1, "
            sql &= "'' AS I1, '' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1 "
            sql &= "From V_ColorRegistr_Digital "
            sql &= "Where BuyerCode = '" & pBuyer & "' "
            '
            If DFKEY.Text <> "" Then
                '使用A1~G1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "YKKColorCode/CustomerColor+CustomerColorCode/Customer+CustomerCode/TapeType/No/Name/Sts/ApplyDate/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And YKKColorCode+CustomerColor+CustomerColorCode+Customer+CustomerCode+TapeType+No+Name+Sts+ApplyDate Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And YKKColorCode+CustomerColor+CustomerColorCode+Customer+CustomerCode+TapeType+No+Name+Sts+ApplyDate Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by No DESC,YKKColorCode, CustomerColor, CustomerColorCode, Customer, CustomerCode, TapeType "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '--FAS
        If AtFAS.Checked = True Then
            sql = "SELECT TOP 300 "
            sql = sql + "'DTMW' as DTMW, "
            sql = sql + "'' as DTMWURL, "
            sql = sql + "'SLDC' as SLDC, "
            sql = sql + "'' as SLDCURL, "
            sql = sql + "'IRW' as IRW, "
            sql = sql + "'' as IRWURL, "
            '
            sql &= "Season AS A1, "
            sql &= "Color1 AS B1, "
            sql &= "Color2 AS C1, "
            sql &= "Green AS D1, "
            sql &= "YColor AS E1, "
            sql &= "'' AS F1, '' AS G1, '' AS H1, '' AS I1, '' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1 "
            sql &= "From M_ColorConvert "
            If HBuyerCode.Text = "FALL-000001" Then
                sql &= "Where Buyer IN ('FALL-000001','FALL-000016') "
            Else
                sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~E1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "Season/COLOR1/COLOR2/Green/YColor/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And Season+COLOR1+COLOR2+Green+YColor Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And Season+COLOR1+COLOR2+Green+YColor Like '%" & DKEY2.Text & "%' "
            '
            sql &= " Group by Season, COLOR1, COLOR2, Green, YColor "
            sql &= " Order by Season Desc, COLOR1, COLOR2, Green, YColor "

            Dim dt_Item As DataTable = uDataBase.GetDataTable(sql)
            If dt_Item.Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = dt_Item
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '--SALES
        If AtSales.Checked = True Then
            sql = "SELECT top 300 "
            sql = sql + "'DTMW' as DTMW, "
            sql = sql + "'' as DTMWURL, "
            sql = sql + "'SLDC' as SLDC, "
            sql = sql + "'' as SLDCURL, "
            sql = sql + "'IRW' as IRW, "
            sql = sql + "'' as IRWURL, "
            '
            sql &= "Season AS A1, "
            sql &= "SeasonYear AS B1, "
            sql &= "Color AS C1, "
            sql &= "BuyerColor AS D1, "
            sql &= "QTY AS E1, "
            sql &= "'' AS F1, '' AS G1, '' AS H1, '' AS I1, '' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1 "
            sql &= "From V_OPReportData_DigitalColor "
            sql &= "Where Buyer = '" & pBuyer & "' "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "Season/SeasonYear/Color/BuerColor/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And Season+convert(varchar,seasonyear)+Color+BuyerColor Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And Season+convert(varchar,seasonyear)+Color+BuyerColor Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by SeasonYear Desc, Season Desc, QTY Desc, Color, BuyerColor  "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '=========================================================================
        'BUYER別
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN G1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(E1)) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(G1)) "
                sql = sql + " END AS SLDCURL, "
                sql &= "'IRW' AS IRW, "
                sql &= "'http://10.245.0.3/ISOS/ISOS2IRW.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pSlider1=' + LTRIM(RTRIM(C1)) as IRWURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用Color
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'SLDC' as SLDC, "
                sql = sql + "'' as SLDCURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACECOLOR" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'ADIDAS  
        '=========================================================================
        If pBuyer = "000001" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN G1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(F1)) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(G1)) "
                sql = sql + " END AS SLDCURL, "
                sql &= "'IRW' AS IRW, "
                sql &= "'http://10.245.0.3/ISOS/ISOS2IRW.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pSlider1=' + LTRIM(RTRIM(C1)) + LTRIM(RTRIM(D1)) as IRWURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用Color
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'SLDC' as SLDC, "
                sql = sql + "'' as SLDCURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACECOLOR" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN F1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(E1)) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(F1)) "
                sql = sql + " END AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    If InStr("5N18/6N18/7N18", DKEY1.Text) > 0 Then
                        sql &= " And ( "
                        sql &= "    A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "5N18" & "%' "
                        sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "6N18" & "%' "
                        sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "7N18" & "%' "
                        sql &= " ) "
                        sql &= " And B1 NOT IN ('N5N18','N6N18','N7N18') "
                    Else
                        If InStr("N5N18/N6N18/N7N18", DKEY1.Text) > 0 Then
                            sql &= " And ( "
                            sql &= "    A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "N5N18" & "%' "
                            sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "N6N18" & "%' "
                            sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "N7N18" & "%' "
                            sql &= " ) "
                        Else
                            'sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                            sql &= " And A1+'/'+B1+'/'+C1+'/'+D1+'/'+E1+'/'+F1+'/'+G1+'/'+H1+'/'+I1+'/'+J1+'/'+K1+'/'+L1+'/'+M1+'/'+N1+'/'+O1+'/'+P1+'/'+Q1+'/'+R1+'/'+S1+'/'+T1+'/'+U1+'/'+V1+'/'+W1+'/'+X1+'/'+Y1+'/'+Z1 Like '%" & DKEY1.Text & "%' "
                        End If
                    End If
                End If

                If DKEY2.Text <> "" Then
                    If InStr("5N18/6N18/7N18", DKEY1.Text) > 0 Then
                        sql &= " And ( "
                        sql &= "    A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "5N18" & "%' "
                        sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "6N18" & "%' "
                        sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "7N18" & "%' "
                        sql &= " ) "
                        sql &= " And B1 NOT IN ('N5N18','N6N18','N7N18') "
                    Else
                        If InStr("N5N18/N6N18/N7N18", DKEY1.Text) > 0 Then
                            sql &= " And ( "
                            sql &= "    A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "N5N18" & "%' "
                            sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "N6N18" & "%' "
                            sql &= " Or A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & "N7N18" & "%' "
                            sql &= " ) "
                        Else
                            'sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                            sql &= " And A1+'/'+B1+'/'+C1+'/'+D1+'/'+E1+'/'+F1+'/'+G1+'/'+H1+'/'+I1+'/'+J1+'/'+K1+'/'+L1+'/'+M1+'/'+N1+'/'+O1+'/'+P1+'/'+Q1+'/'+R1+'/'+S1+'/'+T1+'/'+U1+'/'+V1+'/'+W1+'/'+X1+'/'+Y1+'/'+Z1 Like '%" & DKEY2.Text & "%' "
                        End If
                    End If
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用Color
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'SLDC' as SLDC, "
                sql = sql + "'' as SLDCURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACECOLOR" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN F1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(E1)) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(F1)) "
                sql = sql + " END AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用Color
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'SLDC' as SLDC, "
                sql = sql + "'' as SLDCURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACECOLOR" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN F1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(E1)) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(F1)) "
                sql = sql + " END AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用Color
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'SLDC' as SLDC, "
                sql = sql + "'' as SLDCURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACECOLOR" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'UNDER ARMOUR
        '=========================================================================
        If pBuyer = "TW0371" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN H1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(D1)) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(H1)) "
                sql = sql + " END AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用Color
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'SLDC' as SLDC, "
                sql = sql + "'' as SLDCURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACECOLOR" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'HELLY HANSEN
        '=========================================================================
        If pBuyer = "000098" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(F1)) "
                sql = sql + " AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'BURTON
        '=========================================================================
        If pBuyer = "000068" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(REPLACE(F1,'#',''))) "
                sql = sql + " AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'PUMA
        '=========================================================================
        If pBuyer = "000151" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN G1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(REPLACE(F1,'#',''))) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(REPLACE(G1,'#',''))) "
                sql = sql + " END AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        ' HERSCHEL
        '=========================================================================
        If pBuyer = "TW5068" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "CASE WHEN G1='' THEN "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(REPLACE(F1,'#',''))) "
                sql = sql + " ELSE "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(REPLACE(F1,'#',''))) "
                sql = sql + " END AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & "ALL" & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        ' VERA BRADLEY
        '=========================================================================
        If pBuyer = "TW0655" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'' as SLDC, "
                sql = sql + "'' AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & "ALL" & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        '
        '=========================================================================
        ' PATAGONIA
        '=========================================================================
        If pBuyer = "000141" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'' as SLDC, "
                sql = sql + "'' AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        '
        '=========================================================================
        ' ARCTERYX
        '=========================================================================
        If pBuyer = "000053" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                '
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(D1)) "
                sql = sql + " AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by D1,E1,A1,B1,C1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        '=========================================================================
        ' SALOMON
        '=========================================================================
        If pBuyer = "000055" Then
            '
            '--BUYERColor
            If AtBuyer.Checked = True Then
                '
                sql = "SELECT TOP 300 "
                sql = sql + "'DTMW' as DTMW, "
                sql = sql + "'' as DTMWURL, "
                sql = sql + "'IRW' as IRW, "
                sql = sql + "'' as IRWURL, "
                '
                sql = sql + "'兼' as SLDC, "
                sql = sql + "'http://10.245.1.6/WorkFlowSub/ISOSGetColorFA300.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pColor=' + LTRIM(RTRIM(D1)) "
                sql = sql + " AS SLDCURL, "
                '
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                sql &= "And Active = '0' "
                sql &= "And A1 = '" & DColorType.SelectedValue & "' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~I1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                End If

                If DKEY2.Text <> "" Then
                    sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                End If
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "

                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:


    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Get Search Field
    '**
    '*****************************************************************
    Public Function GetSearchField(ByVal pCmd As String, ByVal pFieldStr As String, ByVal pDataStr As String) As String
        Dim RtnString As String = ""
        Dim str As String
        Dim i As Integer
        Dim FieldNames(), DataNames() As String
        '
        Try
            str = UCase(pCmd)
            FieldNames = pFieldStr.Split("/")
            DataNames = pDataStr.Split("/")
            For i = 0 To DataNames.Length - 1
                If FieldNames(i) <> DataNames(i) Then
                    str = str.Replace(FieldNames(i), DataNames(i))
                End If
            Next
            RtnString = str
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        '=========================================================================
        '共通
        '=========================================================================
        '
        '**開發
        If AtDevlop.Checked = True Then
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Color"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "Customer Color"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "Customer"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "DTMW Sheet"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "DTMW NO."
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "F1" & "<BR>" & "申請者"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "G1" & "<BR>" & "Status"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "H1" & "<BR>" & "申請日"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 8 To 28
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        '**FAS
        If AtFAS.Checked = True Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "Tape"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "Teeth / Slider"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "Other"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "Color"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 5 To 28
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        '--SALES
        If AtSales.Checked = True Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "Year"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "Color"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "Buyer Color"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "Qty"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(4).Text = String.Format("{0:N0}", CDbl(e.Row.Cells(4).Text))
                For i = 5 To 28
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        '=========================================================================
        'BUYER別
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Tape/Puller"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "[T]TWN Color"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "[T]Statsu(P)"
                    tcl(5).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "[T]PBR Color"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "[T]Statsu(PBR)"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "[P]Statsu"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then

                    If e.Row.Cells(0).Text = "TAPE" Then

                        If (InStr(Replace(e.Row.Cells(5).Text, " ", ""), "NOTYET") > 0 Or InStr(Replace(e.Row.Cells(7).Text, " ", ""), "NOTYET") > 0) Then

                            Dim h As HyperLink = e.Row.Cells(26).Controls(0)
                            h.NavigateUrl = "http://10.245.0.3/ISOS/PS_INQ_Color01.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pDTMWColor=" & e.Row.Cells(3).Text
                        Else
                            e.Row.Cells(26).Text = ""
                        End If
                        '
                        For i = 9 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                    Else
                        '
                        For i = 9 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                    '
                    If e.Row.Cells(0).Text = "PULLER" Then e.Row.Cells(27).Visible = False 'SLDC
                    '
                    e.Row.Cells(28).Visible = False
                    If e.Row.Cells(0).Text = "PULLER" Then
                        If InStr(e.Row.Cells(8).Text, "APPROVED") > 0 Then
                            If InStr("SL034/MK045/IT003", UCase(UserID)) > 0 Then
                                e.Row.Cells(28).Visible = True
                            End If
                        End If
                    End If
                End If
            End If
            '
            '**代用COLOR
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Buyer Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "YKK代用色"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "TNF Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Remarks"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(2).ForeColor = Color.Red
                    '
                    For i = 5 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
        End If
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Brand"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "YKKColor"
                        tcl(5).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "PBR*D"
                        tcl(6).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "DT*/VT*/CNT*/CFT*" & "<BR>" & "(with P tape)"
                        tcl(7).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "DT*/VT*/CZT*/D*" & "<BR>" & "(with Recycled tape)"
                        tcl(8).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(9).Text = "J1" & "<BR>" & "NM T.Color"
                        tcl.Add(New TableHeaderCell())
                        tcl(10).Text = "K1" & "<BR>" & "YKK S.Color"
                        tcl.Add(New TableHeaderCell())
                        tcl(11).Text = "L1" & "<BR>" & "YKK S.Color " & "<BR>" & "for EAA/EFA/EX Neon color" & "<BR>" & "(SLS-*** / SLS*** / SLS#****)"
                        tcl.Add(New TableHeaderCell())
                        tcl(12).Text = "M1" & "<BR>" & "P Tape(VSNT)"
                        tcl.Add(New TableHeaderCell())
                        tcl(13).Text = "N1" & "<BR>" & "PBR*D Tape(VSNT)"
                        tcl.Add(New TableHeaderCell())
                        tcl(14).Text = "O1" & "<BR>" & "Clear Teeth(VSNT)"
                        tcl.Add(New TableHeaderCell())
                        tcl(15).Text = "P1" & "<BR>" & "SLS-***(VSNT)"
                        tcl.Add(New TableHeaderCell())
                        tcl(16).Text = "Q1" & "<BR>" & "Remark"
                        'tcl.Add(New TableHeaderCell())
                        'tcl(17).Text = "R1" & "<BR>" & "Remark(for MKT)"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(3).ForeColor = Color.Red
                        e.Row.Cells(8).Width = 110
                        For i = 17 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Puller Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Color Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "核可樣送核日"
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(3).ForeColor = Color.Red
                        For i = 8 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                        '
                        If e.Row.Cells(0).Text = "PULLER" Then
                            If InStr(e.Row.Cells(1).Text, "APPROVAL") > 0 Then
                                If InStr("SL034/MK045/IT003", UCase(UserID)) > 0 Then
                                    e.Row.Cells(28).Visible = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            '
            '**代用COLOR
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "OrderType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "依賴編號"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YKK代用色"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(6).ForeColor = Color.Red
                    '
                    For i = 7 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
        End If
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Labdip Vendor Code" & "<BR>" & "(YKK code)"
                        tcl(4).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "PBR2D/PBR6D"
                        tcl(5).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "NM(transparent tape)" & "<BR>" & "Labdip Vendor Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "SHA Color"
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "DLN Color"
                        tcl.Add(New TableHeaderCell())
                        tcl(9).Text = "J1" & "<BR>" & "Japan Color"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        For i = 10 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'TEETH
                If DColorType.SelectedValue = "TEETH" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "K1" & "<BR>" & "Teeth"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        For i = 3 To 9
                            e.Row.Cells(i).Visible = False
                        Next
                        For i = 11 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
                '
                'SLDFINISH
                If DColorType.SelectedValue = "SLDFINISH" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Labdip Vendor Code" & "<BR>" & "(YKK code)"
                        tcl(4).BackColor = Color.Blue
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        For i = 5 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Puller"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Slider Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Base"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "核可樣送核日"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "核可狀態"
                        tcl(5).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        If InStr(e.Row.Cells(3).Text, "GRIND") > 0 Then e.Row.Cells(3).ForeColor = Color.Blue

                        For i = 7 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
            '
            '**代用COLOR
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "OrderType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "依賴編號"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YKK代用色"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(6).ForeColor = Color.Red
                    '
                    For i = 7 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False 'SLDC
                End If
            End If
            '
        End If
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "State"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "LabDipCode"
                        tcl(4).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "PBR*D"
                        tcl(5).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "CNT8/10"
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "VT8/10"
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "Remark"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        'e.Row.Cells(3).ForeColor = Color.Red
                        For i = 9 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Puller"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Slider Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Base"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "核可狀態"
                        tcl(5).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        For i = 7 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
            '
            '**代用COLOR
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "OrderType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "依賴編號"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YKK代用色"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(6).ForeColor = Color.Red
                    '
                    For i = 7 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False 'SLDC
                End If
            End If
            '
        End If
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "TWN Color"
                        tcl(4).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "PBR Color"
                        tcl(5).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "HK Color"
                        tcl(6).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "SHA Color"
                        tcl(7).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "SZN Color"
                        tcl(8).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(9).Text = "J1" & "<BR>" & "VNM Color"
                        tcl(9).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(10).Text = "K1" & "<BR>" & "JP Color"
                        tcl(10).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(11).Text = "L1" & "<BR>" & "Korea Color"
                        tcl(11).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(12).Text = "M1" & "<BR>" & "Dalian Color"
                        tcl(12).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(13).Text = "N1" & "<BR>" & "Soft Clear Zipper NM"
                        tcl(13).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(14).Text = "O1" & "<BR>" & "Remark"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(3).ForeColor = Color.Blue
                        e.Row.Cells(4).ForeColor = Color.Red
                        For i = 15 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "核可狀態"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Name"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Puller Name"
                        tcl(4).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Puller Code"
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(3).ForeColor = Color.Blue
                        e.Row.Cells(5).ForeColor = Color.Red
                        For i = 7 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
            '
            '**代用COLOR
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "OrderType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "依賴編號"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YKK代用色"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(6).ForeColor = Color.Red
                    '
                    For i = 7 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False 'SLDC
                End If
            End If
            '
        End If
        '=========================================================================
        'UNDER ARMOUR
        '=========================================================================
        If pBuyer = "TW0371" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "SMS ONLY"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Code"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "E/EF/EAA/EFA"
                        tcl(4).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "EFR"
                        tcl(5).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "NEON/GENERAL" & "<BR>" & "(BULK-Y COLOR)"
                        tcl(6).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "PBR"
                        tcl(7).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "Remarks"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(3).ForeColor = Color.Red
                        If InStr(e.Row.Cells(6).Text, "NEON") > 0 Then
                            e.Row.Cells(3).ForeColor = Color.Blue
                            e.Row.Cells(6).ForeColor = Color.Blue
                        End If

                        For i = 9 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Puller Name"
                        tcl(1).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Rubber Puller Code"
                        tcl(2).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Rubber Puller Name"
                        tcl(3).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red

                        If InStr(e.Row.Cells(3).Text, "ONLY") > 0 Then e.Row.Cells(3).ForeColor = Color.Blue

                        For i = 6 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
            '
            '**代用COLOR
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "OrderType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "依賴編號"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YKK代用色"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(6).ForeColor = Color.Red
                    '
                    For i = 7 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False 'SLDC
                End If
            End If
            '
        End If
        '=========================================================================
        'HELLY HANSEN
        '=========================================================================
        If pBuyer = "000098" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "F1" & "<BR>" & "Color Code"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "H1" & "<BR>" & "Remark"
                        tcl(4).BackColor = Color.Blue
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(3).Visible = False
                        e.Row.Cells(4).Visible = False
                        e.Row.Cells(5).ForeColor = Color.Red
                        e.Row.Cells(6).Visible = False
                        For i = 8 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Puller"
                        tcl(2).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Puller Code"
                        tcl(3).BackColor = Color.Green
                        '
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Color"
                        tcl(4).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Finish Color"
                        tcl(5).BackColor = Color.Blue
                        '
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(3).ForeColor = Color.Red

                        For i = 7 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
            '
        End If
        '=========================================================================
        'BURTON
        '=========================================================================
        If pBuyer = "000068" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Color Name"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Pantone"
                        tcl(4).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "YKK"
                        tcl(5).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "DT2"
                        tcl(6).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "VT10"
                        tcl(7).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "CNT10"
                        tcl(8).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(9).Text = "J1" & "<BR>" & "SHA"
                        tcl(9).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(10).Text = "K1" & "<BR>" & "Remark"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(1).Text = Replace(e.Row.Cells(1).Text, Chr(10), "<BR>")
                        e.Row.Cells(5).Text = Replace(e.Row.Cells(5).Text, "#", "")
                        e.Row.Cells(6).Text = Replace(e.Row.Cells(6).Text, "#", "")
                        e.Row.Cells(7).Text = Replace(e.Row.Cells(7).Text, "#", "")
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "#", "")
                        e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, "#", "")

                        e.Row.Cells(5).ForeColor = Color.Red
                        For i = 11 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Puller"
                        tcl(2).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Puller Code"
                        tcl(3).BackColor = Color.Green
                        '
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Color"
                        tcl(4).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Burton Name"
                        tcl(5).BackColor = Color.Blue
                        '
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(3).ForeColor = Color.Red
                        e.Row.Cells(4).Text = Replace(e.Row.Cells(4).Text, "#", "")

                        For i = 7 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
        End If
        '=========================================================================
        'PUMA
        '=========================================================================
        If pBuyer = "000151" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Type"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Pantone"
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "YKK"
                        tcl(5).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "PBR"
                        tcl(6).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "P.tape Approved"
                        tcl(7).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "VT8"
                        tcl(8).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(9).Text = "J1" & "<BR>" & "VT9"
                        tcl(9).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(10).Text = "K1" & "<BR>" & "VT10"
                        tcl(10).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(11).Text = "L1" & "<BR>" & "CNT8/CFT8"
                        tcl(11).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(12).Text = "M1" & "<BR>" & "CNT9/CFT9"
                        tcl(12).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(13).Text = "N1" & "<BR>" & "CNT10/CFT10"
                        tcl(13).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(14).Text = "O1" & "<BR>" & "CNDT1/CFDT1"
                        tcl(14).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(15).Text = "P1" & "<BR>" & "CNDT2/CFDT2"
                        tcl(15).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(16).Text = "Q1" & "<BR>" & "Remark"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(5).ForeColor = Color.Red
                        e.Row.Cells(6).ForeColor = Color.Red

                        For i = 17 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Color Cat."
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Status"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Puller"
                        tcl(2).BackColor = Color.Blue
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Puller Code"
                        tcl(3).BackColor = Color.Blue
                        '
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Color"
                        tcl(4).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Finish color"
                        tcl(5).BackColor = Color.Green
                        '
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Remark"
                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(3).ForeColor = Color.Red

                        For i = 7 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
        End If
        '=========================================================================
        'HERSCHEL
        '=========================================================================
        If pBuyer = "TW5068" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then

                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Pantone"
                    tcl.Add(New TableHeaderCell())

                    tcl(4).Text = "E1" & "<BR>" & "Tape (for SAMPLES)"
                    tcl(4).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Tape"
                    tcl(5).BackColor = Color.Red

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Slider (for SAMPLES)"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Slider"
                    tcl(7).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "RCPU (for SAMPLES)"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "RCPU"
                    tcl(9).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(5).ForeColor = Color.Red
                    e.Row.Cells(7).ForeColor = Color.Red
                    e.Row.Cells(9).ForeColor = Color.Red
                    '
                    '調整特定欄位寬度
                    e.Row.Cells(10).Width = 300
                    '
                    For i = 11 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(28).Visible = False
                End If
            End If
        End If
        '=========================================================================
        'VERA BRADLEY
        '=========================================================================
        If pBuyer = "TW0655" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then

                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Color Name"

                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Colo Type"
                    tcl(3).BackColor = Color.Red

                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Photo Color"
                    tcl(4).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Production Color"
                    tcl(5).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YOC Photo Color"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "YOC Production Color"
                    tcl(7).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    'e.Row.Cells(3).Font.Bold = True

                    e.Row.Cells(5).ForeColor = Color.Red
                    e.Row.Cells(7).ForeColor = Color.Red
                    '
                    '調整特定欄位寬度
                    e.Row.Cells(8).Width = 300
                    '
                    For i = 9 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                    e.Row.Cells(28).Visible = False
                End If
            End If
        End If
        '=========================================================================
        'PATAGONIA
        '=========================================================================
        If pBuyer = "000141" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then

                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Swatch Color"

                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "YKK Color Code"
                    tcl(3).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "VT8/10" & "<BR>" & "(CC-DTM2)"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "CNT/CFT" & "<BR>" & "(CC-DTM1)"

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Remark"


                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    'e.Row.Cells(3).Font.Bold = True

                    'e.Row.Cells(5).ForeColor = Color.Red
                    ''
                    ''調整特定欄位寬度
                    'e.Row.Cells(6).Width = 300
                    '
                    For i = 7 To 26
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                    e.Row.Cells(28).Visible = False
                End If
            End If
        End If

        '=========================================================================
        'ARCTERYX
        '=========================================================================
        If pBuyer = "000053" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Type"
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Standard"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "tape color (Approved)"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "ESM/ERP slider color  (Approved)"
                        tcl(4).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Korea color"
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "Japan color"
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "Canada color"
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "Remark"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(1).Text = Replace(e.Row.Cells(1).Text, Chr(10), "<BR>")
                        e.Row.Cells(5).Text = Replace(e.Row.Cells(5).Text, "#", "")
                        e.Row.Cells(6).Text = Replace(e.Row.Cells(6).Text, "#", "")
                        e.Row.Cells(7).Text = Replace(e.Row.Cells(7).Text, "#", "")
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "#", "")
                        e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, "#", "")

                        For i = 9 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Type"
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Color"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Standard"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Slider Code"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Puller Code"
                        tcl(4).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Remark"

                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        For i = 6 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
        End If

        '=========================================================================
        'SALOMON
        '=========================================================================
        If pBuyer = "000055" Then
            '
            '**BUYERColor
            If AtBuyer.Checked = True Then
                'TAPE
                If DColorType.SelectedValue = "TAPE" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Type"
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Season"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Specification"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Stander Tape"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "CNT10"
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "CNT8"
                        tcl.Add(New TableHeaderCell())
                        tcl(6).Text = "G1" & "<BR>" & "VT8"
                        tcl.Add(New TableHeaderCell())
                        tcl(7).Text = "H1" & "<BR>" & "VT10"
                        tcl.Add(New TableHeaderCell())
                        tcl(8).Text = "I1" & "<BR>" & "Remark"
                    End If
                    '
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        e.Row.Cells(1).Text = Replace(e.Row.Cells(1).Text, Chr(10), "<BR>")
                        e.Row.Cells(3).Text = Replace(e.Row.Cells(3).Text, "#", "")
                        e.Row.Cells(4).Text = Replace(e.Row.Cells(4).Text, "#", "")
                        e.Row.Cells(5).Text = Replace(e.Row.Cells(5).Text, "#", "")
                        e.Row.Cells(6).Text = Replace(e.Row.Cells(6).Text, "#", "")
                        e.Row.Cells(7).Text = Replace(e.Row.Cells(7).Text, "#", "")

                        For i = 9 To 26
                            e.Row.Cells(i).Visible = False
                        Next
                        e.Row.Cells(28).Visible = False
                    End If
                End If
                '
                'PULLER
                If DColorType.SelectedValue = "PULLER" Then
                    'Header
                    If (e.Row.RowType = DataControlRowType.Header) Then
                        Dim tcl As TableCellCollection = e.Row.Cells
                        tcl.Clear() '清除自动生成的表头
                        '添加新的表头第一行
                        tcl.Add(New TableHeaderCell())
                        tcl(0).Text = "A1" & "<BR>" & "Type"
                        tcl.Add(New TableHeaderCell())
                        tcl(1).Text = "B1" & "<BR>" & "Color Name"
                        tcl.Add(New TableHeaderCell())
                        tcl(2).Text = "C1" & "<BR>" & "Color Standard"
                        tcl.Add(New TableHeaderCell())
                        tcl(3).Text = "D1" & "<BR>" & "Slider Code"
                        tcl(3).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(4).Text = "E1" & "<BR>" & "Puller Code"
                        tcl(4).BackColor = Color.Green
                        tcl.Add(New TableHeaderCell())
                        tcl(5).Text = "F1" & "<BR>" & "Remark"

                    End If
                    'DataRow
                    If (e.Row.RowType = DataControlRowType.DataRow) Then
                        For i = 6 To 28
                            e.Row.Cells(i).Visible = False
                        Next
                    End If
                End If
            End If
        End If
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
    '
    '*****************************************************************
    '**
    '**     CHECKBOX CHANGE
    '**
    '*****************************************************************
    Protected Sub AtBuyer_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtBuyer.CheckedChanged
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        AtDevlop.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtDevlop.Checked = False Then
            AtBuyer.Checked = True
        End If
    End Sub

    Protected Sub AtFAS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtFAS.CheckedChanged
        AtBuyer.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        AtDevlop.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtDevlop.Checked = False Then
            AtFAS.Checked = True
        End If
    End Sub

    Protected Sub AtReplace_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtReplace.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtSales.Checked = False
        AtDevlop.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtDevlop.Checked = False Then
            AtReplace.Checked = True
        End If
    End Sub

    Protected Sub AtSales_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtSales.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtDevlop.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtDevlop.Checked = False Then
            AtSales.Checked = True
        End If
    End Sub

    Protected Sub AtDevlop_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtDevlop.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtDevlop.Checked = False Then
            AtDevlop.Checked = True
        End If
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=ColorList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

End Class
