Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_ListPrice01
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
    Dim pBuyer As String             'Buyer
    Dim UserID As String            'UserID
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
        Response.Cookies("PGM").Value = "PS_INQ_ListPrice01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        DINQLISTPRICEFile.Style("left") = -500 & "px"
        DINQLISTPRICEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQLISTPRICE_" & UserID & ".xlsm"
        '
        'label5 備註說明
        Label5.Style("left") = -500 & "px"
        If pBuyer = "000013" Then
            Label5.Text = "Unit: USD$/per 100pcs"
            Label5.Style("left") = 712 & "px"
        End If
        If pBuyer = "000003" Then
            Label5.Text = "Zipper Unit: USD$/per 100pcs"
            Label5.Style("left") = 712 & "px"
        End If
        If pBuyer = "000098" Or pBuyer = "000068" Then
            Label5.Text = "Unit: TWD$/per (Zipper=100pcs / Slider=1000pcs)"
            Label5.Style("left") = 712 & "px"
        End If
        If pBuyer = "000053" Then
            Label5.Text = "Zipper unit: USD$/per 100pcs"
            Label5.Style("left") = 712 & "px"
        End If
        '
        '動作按鈕設定
        BSPInq.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('INQLISTPRICEExcel')}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
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
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            If pBuyer = "000098" Then DBUYER.Text = "HELLY HANSEN"
            If pBuyer = "000068" Then DBUYER.Text = "BURTON"
            If pBuyer = "TW5068" Then DBUYER.Text = "HERSCHEL"
            If pBuyer = "TW0655" Then DBUYER.Text = "VERA BRADLEY"
            If pBuyer = "000141" Then DBUYER.Text = "PATAGONIA"
            If pBuyer = "000053" Then DBUYER.Text = "ARCTERYX"
            If pBuyer = "000055" Then DBUYER.Text = "SALOMON"
            If pBuyer = "TW0020" Then DBUYER.Text = "TUMI"
            '
            If pBuyer <> "000021" And pBuyer <> "000003" And pBuyer <> "000141" And pBuyer <> "000053" And pBuyer <> "000055" Then DSeason.Style("left") = -500 & "px"
            '
            If pBuyer <> "TW0371" Then DItemType.Style("left") = -500 & "px"
            '
            HBuyerCode.Text = pBuyer
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        On Error GoTo LBL_Error

        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '
        '=========================================================================
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + A1 + '&pItemName=' + Replace(P1,'+','*') as URL "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If DSeason.SelectedValue = "999999" Then
                sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            Else
                sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/O1/P1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + A1 + '&pItemName=' + Replace(H1,'+','*') as URL "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            'sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/O1/P1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(Replace(V1,'#',''),'~','') + '&pYCode=' + U1 as URL "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/H1/I1/J1/K1/L1/M1/N1/O1/P1/S1/T1/U1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + E1 + '&pItemName=' + Replace(Replace(U1,'#',''),'~','') + '&pYCode=' + T1 as URL "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If DSeason.SelectedValue = "999999" Then
                sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            Else
                sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        '=========================================================================
        'UA
        '=========================================================================
        If pBuyer = "TW0371" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            If DItemType.Text = "APP" Then
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(Replace(O1,'#',''),'~','') + '&pYCode=' + N1 as URL "
            Else
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + F1 + '&pItemName=' + Replace(Replace(Q1,'#',''),'~','') + '&pYCode=' + P1 as URL "
            End If
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            sql &= "And A1 =  '" & DItemType.Text & "' "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by A1,B1,C1,D1 DESC,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
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
        'HELLY HANSEN
        '=========================================================================
        If pBuyer = "000098" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + G1 + '&pItemName=' + Replace(Replace(H1,'#',''),'~','') + '&pYCode=' + G1 as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            sql &= "And J1 > '0' "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        '=========================================================================
        'BURTON
        '=========================================================================
        If pBuyer = "000068" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + G1 + '&pItemName=' + Replace(Replace(H1,'#',''),'~','') + '&pYCode=' + G1 as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            sql &= "And J1 > '0' "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        '=========================================================================
        'HERSCHEL
        '=========================================================================
        If pBuyer = "TW5068" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + Replace(Replace(Replace(G1,'#',''),'~',''),'*','') + '&pYCode=' as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        '=========================================================================
        'VERA BRADLEY
        '=========================================================================
        If pBuyer = "TW0655" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(O1,'+','*') + '&pYCode=' + N1 as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
        ''=========================================================================
        ''PATAGONIA
        ''=========================================================================
        If pBuyer = "000141" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + C1 + '&pItemName=' + Replace(O1,'+','*') + '&pYCode=' + N1 as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If DSeason.SelectedValue = "999999" Then
                sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            Else
                sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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

        ''=========================================================================
        ''ARCTERYX
        ''=========================================================================
        If pBuyer = "000053" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(G1,'+','*') + '&pYCode=' + F1 as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If DSeason.SelectedValue = "999999" Then
                sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            Else
                sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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


        ''=========================================================================
        ''SALOMON
        ''=========================================================================
        If pBuyer = "000055" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(F1,'+','*') + '&pYCode=' as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If DSeason.SelectedValue = "999999" Then
                sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            Else
                sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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

        ''=========================================================================
        ''TUMI
        ''=========================================================================
        If pBuyer = "TW0020" Then
            '
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
            '*
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=&pItemName=' + Replace(G1,'+','*') + '&pYCode=' + F1 as URL "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If DSeason.SelectedValue = "999999" Then
                sql &= "And DataType IN ('" & "BUYERPRICE" & "', '" & "BUYERPRICE-B" & "') "
            Else
                sql &= "And DataType IN ('" & DSeason.SelectedValue & "') "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
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
        '
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
                    If str <> UCase(pCmd) Then Exit For
                End If
            Next
            RtnString = str
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=CustomerOrder.xls")     '程式別不同
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
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Cat."
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Type"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Web Code"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Description"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Article"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Base length (in inch, for zipper only)"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Price for basic length (USD/100pcs)"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "Price for 1inch up ( USD/100pcs)"
                tcl(9).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "YKK/Web Code"
                tcl(10).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "Article" & "<BR>" & "(Puller:Tnf Black僅供參考)"
                tcl(11).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                If InStr(e.Row.Cells(16).Text, "[1]") > 0 Then
                    e.Row.Cells(15).ForeColor = Color.Red
                    e.Row.Cells(16).Text = Replace(e.Row.Cells(16).Text, "[1]", "")
                Else
                    If InStr(e.Row.Cells(16).Text, "[]") > 0 Then
                        e.Row.Cells(15).ForeColor = Color.Blue
                        e.Row.Cells(16).Text = Replace(e.Row.Cells(16).Text, "[]", "")
                    End If
                End If

                For i = 10 To 14
                    e.Row.Cells(i).Visible = False
                Next
                '
                For i = 17 To 26
                    e.Row.Cells(i).Visible = False
                Next

            End If
        End If
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Brand"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Category"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Item Status"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "End Season"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Kids safe"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "PLM#"
                tcl(7).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "YKK ITEM"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "Base length"
                tcl(9).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "Base (US$/PC)"
                tcl(10).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "add.1 inch (US$/PC)"
                tcl(11).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "Bulk L/T"
                tcl.Add(New TableHeaderCell())
                tcl(13).Text = "M1" & "<BR>" & "Sample L/T"
                tcl.Add(New TableHeaderCell())
                tcl(14).Text = "N1" & "<BR>" & "YKK/PLM"
                tcl(14).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(15).Text = "O1" & "<BR>" & "YKK ITEM" & "<BR>" & "(Puller:095A僅供參考)"
                tcl(15).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, Chr(10), "<br>")
                e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, Chr(10), "<br>")

                If InStr(e.Row.Cells(15).Text, "[1]") > 0 Then
                    e.Row.Cells(14).ForeColor = Color.Red
                    e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[1]", "")
                Else
                    If InStr(e.Row.Cells(15).Text, "[]") > 0 Then
                        e.Row.Cells(14).ForeColor = Color.Blue
                        e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[]", "")
                    End If
                End If
                For i = 16 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Kids Safe"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "PCX"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "IM Code"
                tcl(4).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Status"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Description"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Base" & "<BR>" & "Size(inch)"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Base"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "Add 1 inch"
                tcl(9).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "YPGOLD" & "<BR>" & "(OQG)"
                tcl(10).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "YKKSHB" & "<BR>" & "(VKL)"
                tcl(11).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "SHINY ICESIL" & "<BR>" & "(C5V)"
                tcl(12).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(13).Text = "M1" & "<BR>" & "YKKH6N" & "<BR>" & "(H6N)"
                tcl(13).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(14).Text = "N1" & "<BR>" & "01M" & "<BR>" & "(H1)"
                tcl(14).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(15).Text = "O1" & "<BR>" & "OP-O(95K)" & "<BR>" & "OP-H6(95H)"
                tcl(15).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(16).Text = "P1" & "<BR>" & "Sample L/T"
                tcl.Add(New TableHeaderCell())
                tcl(17).Text = "Q1" & "<BR>" & "Bulk L/T"
                tcl.Add(New TableHeaderCell())
                tcl(18).Text = "R1" & "<BR>" & "MCO_zipper"
                tcl.Add(New TableHeaderCell())
                tcl(19).Text = "S1" & "<BR>" & "Remark"
                tcl.Add(New TableHeaderCell())
                tcl(20).Text = "T1" & "<BR>" & "YKK/IM Code"
                tcl(20).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(21).Text = "U1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:00A僅供參考)"
                tcl(21).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(8).ForeColor = Color.Red
                e.Row.Cells(9).ForeColor = Color.Red
                e.Row.Cells(10).ForeColor = Color.Red

                For i = 1 To 24
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                Next
                If InStr(e.Row.Cells(22).Text, "[1]") > 0 Then
                    e.Row.Cells(21).ForeColor = Color.Red
                    e.Row.Cells(22).Text = Replace(e.Row.Cells(22).Text, "[1]", "")
                Else
                    If InStr(e.Row.Cells(22).Text, "[]") > 0 Then
                        e.Row.Cells(21).ForeColor = Color.Blue
                        e.Row.Cells(22).Text = Replace(e.Row.Cells(22).Text, "[]", "")
                    End If
                End If
                '
                e.Row.Cells(7).Visible = False
                'For i = 17 To 18
                '    e.Row.Cells(i).Visible = False
                'Next
                For i = 23 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Brand"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Category"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Status"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Kids Safe"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "PDM"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "YKK Dsc"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Description"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Base Length"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "L/T" & "<BR>" & "SPL w/o"
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "L/T" & "<BR>" & "SPL w/"
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "L/T" & "<BR>" & "PROD w/o"
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "L/T" & "<BR>" & "PROD w/"
                tcl.Add(New TableHeaderCell())
                tcl(13).Text = "M1" & "<BR>" & "Unit"
                tcl(13).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(14).Text = "N1" & "<BR>" & "Currency"
                tcl(14).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(15).Text = "O1" & "<BR>" & "PROD Loc."
                tcl(15).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(16).Text = "P1" & "<BR>" & "Base"
                tcl(16).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(17).Text = "Q1" & "<BR>" & "Add 1 inch"
                tcl(17).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(18).Text = "R1" & "<BR>" & "MOQ" & "<BR>" & "SPL/PROD"
                tcl.Add(New TableHeaderCell())
                tcl(19).Text = "S1" & "<BR>" & "Remark"
                tcl.Add(New TableHeaderCell())
                tcl(20).Text = "T1" & "<BR>" & "YKK/PDM"
                tcl(20).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(21).Text = "U1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                tcl(21).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(8).ForeColor = Color.Red
                e.Row.Cells(16).ForeColor = Color.Red
                e.Row.Cells(17).ForeColor = Color.Red

                For i = 1 To 24
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                Next
                If InStr(e.Row.Cells(21).Text, "[1]") > 0 Then
                    e.Row.Cells(20).ForeColor = Color.Red
                    e.Row.Cells(21).Text = Replace(e.Row.Cells(21).Text, "[1]", "")
                Else
                    If InStr(e.Row.Cells(21).Text, "[]") > 0 Then
                        e.Row.Cells(20).ForeColor = Color.Blue
                        e.Row.Cells(21).Text = Replace(e.Row.Cells(21).Text, "[]", "")
                    End If
                End If
                '
                'For i = 3 To 4
                '    e.Row.Cells(i).Visible = False
                'Next
                'For i = 9 To 12
                '    e.Row.Cells(i).Visible = False
                'Next
                For i = 22 To 25
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '=========================================================================
        'UA
        '=========================================================================
        If pBuyer = "TW0371" Then
            ' 
            If DItemType.SelectedValue = "APP" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "Select"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "A1" & "<BR>" & "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "C1" & "<BR>" & "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "D1" & "<BR>" & "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "E1" & "<BR>" & "Description"
                    tcl(5).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "F1" & "<BR>" & "YKK" & "<BR>" & "Description"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "G1" & "<BR>" & "CPSC"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "H1" & "<BR>" & "SMS" & "<BR>" & "Lead Time"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "I1" & "<BR>" & "BULK" & "<BR>" & "Lead Time"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "J1" & "<BR>" & "Length" & "<BR>" & "(Inch)"
                    tcl(10).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "K1" & "<BR>" & "Price"
                    tcl(11).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "L1" & "<BR>" & "1(Inch)" & "<BR>" & "Upcharge"
                    tcl(12).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "M1" & "<BR>" & "Remark"

                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "N1" & "<BR>" & "YKK/ZP"
                    tcl(14).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "O1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                    tcl(15).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then

                    e.Row.Cells(4).ForeColor = Color.Red

                    For i = 1 To 24
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If InStr(e.Row.Cells(15).Text, "[1]") > 0 Then
                        e.Row.Cells(14).ForeColor = Color.Red
                        e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(15).Text, "[]") > 0 Then
                            e.Row.Cells(14).ForeColor = Color.Blue
                            e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[]", "")
                        End If
                    End If
                    '
                    'For i = 7 To 9
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    For i = 16 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DItemType.SelectedValue = "BAG" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "Select"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "A1" & "<BR>" & "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "C1" & "<BR>" & "#"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "D1" & "<BR>" & "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "E1" & "<BR>" & "Description"
                    tcl(5).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "F1" & "<BR>" & "Rebranding" & "<BR>" & "ZP #"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "G1" & "<BR>" & "Rebraning" & "<BR>" & "YKK Code"
                    tcl(7).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "H1" & "<BR>" & "Previous" & "<BR>" & "ZP #"
                    tcl(8).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "I1" & "<BR>" & "Previous" & "<BR>" & "YKK Code"
                    tcl(9).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "J1" & "<BR>" & "SMS" & "<BR>" & "Lead Time"
                    tcl(10).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "K1" & "<BR>" & "BULK" & "<BR>" & "Lead Time"
                    tcl(11).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "L1" & "<BR>" & "Length" & "<BR>" & "(Inch)"
                    tcl(12).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "M1" & "<BR>" & "Price"
                    tcl(13).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "N1" & "<BR>" & "1(Inch)" & "<BR>" & "Upcharge"
                    tcl(14).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "O1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "P1" & "<BR>" & "YKK/PDM"
                    tcl(16).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "Q1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                    tcl(17).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then

                    e.Row.Cells(6).ForeColor = Color.Red

                    For i = 1 To 24
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If InStr(e.Row.Cells(17).Text, "[1]") > 0 Then
                        e.Row.Cells(16).ForeColor = Color.Red
                        e.Row.Cells(17).Text = Replace(e.Row.Cells(17).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(17).Text, "[]") > 0 Then
                            e.Row.Cells(16).ForeColor = Color.Blue
                            e.Row.Cells(17).Text = Replace(e.Row.Cells(17).Text, "[]", "")
                        End If
                    End If
                    '
                    'For i = 10 To 11
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    For i = 18 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(26).Visible = False
                End If
            End If
        End If
        '=========================================================================
        'HELLY HANSEN
        '=========================================================================
        If pBuyer = "000098" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())

                tcl(1).Text = "A1" & "<BR>" & "Brand"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Prod"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Type"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Price" & "<BR>" & "Type"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Price" & "<BR>" & "Version"
                tcl(6).BackColor = Color.Green

                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Code"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Description"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())

                tcl(9).Text = "I1" & "<BR>" & "Length (Inch)"
                tcl(9).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "List Price"
                tcl(10).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "1(Inch)" & "<BR>" & "Upcharge"
                tcl(11).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 1 To 24
                    If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                        e.Row.Cells(7).ForeColor = Color.Red
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                    Else
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                    End If
                    '
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    '
                    Select Case i
                        Case 9 To 11
                            e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,###.#")
                            e.Row.Cells(i).HorizontalAlign = HorizontalAlign.Right
                        Case Else
                    End Select
                Next
                '
                For i = 12 To 25
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '=========================================================================
        'BURTON
        '=========================================================================
        If pBuyer = "000068" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())

                tcl(1).Text = "A1" & "<BR>" & "Brand"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Prod"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Type"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Price" & "<BR>" & "Type"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Price" & "<BR>" & "Version"
                tcl(6).BackColor = Color.Green

                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Code"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Description"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())

                tcl(9).Text = "I1" & "<BR>" & "Length (Inch)"
                tcl(9).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "List Price"
                tcl(10).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "1(Inch)" & "<BR>" & "Upcharge"
                tcl(11).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 1 To 24
                    If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                        e.Row.Cells(7).ForeColor = Color.Red
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                    Else
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                    End If
                    '
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    '
                    Select Case i
                        Case 9 To 11
                            e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,###.#")
                            e.Row.Cells(i).HorizontalAlign = HorizontalAlign.Right
                        Case Else
                    End Select
                Next
                '
                For i = 12 To 25
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '=========================================================================
        'HERSCHEL
        '=========================================================================
        If pBuyer = "TW5068" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())

                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Category"
                tcl(2).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Part #"
                tcl(3).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Size"
                tcl(4).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Matereial Spec"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Color Name"
                tcl(6).BackColor = Color.Green
                '
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Supplier Item"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Unit Of Measure"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "FOB USD"
                tcl(9).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "MOQ"
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "Lead Time"
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "Remark"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 1 To 24
                    If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                        e.Row.Cells(7).ForeColor = Color.Red
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                    Else
                        e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                    End If
                    '
                    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    '
                    Select Case i
                        Case 9 To 9
                            e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###.######")
                            e.Row.Cells(i).HorizontalAlign = HorizontalAlign.Right
                        Case Else
                    End Select
                Next
                '

                For i = 13 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        '=========================================================================
        'VERA BRADLEY
        '=========================================================================
        If pBuyer = "TW0655" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Cat."
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Class"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Code"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Material Name"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "YKK NAME"
                tcl(6).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Length(inch)"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Price(USD)"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "Unit"
                tcl(9).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "1 inch(USD)"
                tcl(10).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "1inch Unit"
                tcl(11).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "Remark"
                tcl.Add(New TableHeaderCell())
                'tcl(13).Text = "M1" & "<BR>" & "SPPrice"
                'tcl.Add(New TableHeaderCell())
                tcl(13).Text = "N1" & "<BR>" & "YKK/Web Code"
                tcl(13).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(14).Text = "O1" & "<BR>" & "YKK/Web Name"
                tcl(14).BackColor = Color.Red
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                If InStr(e.Row.Cells(15).Text, "[1]") > 0 Then
                    e.Row.Cells(14).ForeColor = Color.Red
                    e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[1]", "")
                Else
                    If InStr(e.Row.Cells(15).Text, "[]") > 0 Then
                        e.Row.Cells(14).ForeColor = Color.Blue
                        e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[]", "")
                    End If
                End If
                '
                If e.Row.Cells(13).Text = "1" Then
                    For i = 7 To 11
                        e.Row.Cells(i).ForeColor = Color.Blue
                    Next
                End If
                '
                e.Row.Cells(13).Visible = False
                For i = 16 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If

        '=========================================================================
        'PATAGONIA
        '=========================================================================
        If pBuyer = "000141" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Trim Type"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "RM#"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Description (RM)"
                tcl(4).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "YKK Item Code"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Length (in.)"
                tcl(6).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Base length (USD/100pcs)"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "1i Up (USD/100pcs)"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "BUNDLE PRICE($, USD)"
                tcl(9).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "UOM"
                tcl(10).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "Cost per" & "<BR>" & "slider(USD/100pcs)" & "<BR>" & "chain(USD/100yards)" & "<BR>" & "snaps(USD/1,000pcs)"
                tcl(11).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "UNIT PRICE($, USD)"
                tcl(12).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(13).Text = "M1" & "<BR>" & "Remark"
                tcl.Add(New TableHeaderCell())
                'tcl(14).Text = "N1" & "<BR>" & "Remark"
                'tcl.Add(New TableHeaderCell())
                'tcl(15).Text = "O1" & "<BR>" & "UNIT PRICE($, USD)"
                'tcl(15).BackColor = Color.Red
                'tcl.Add(New TableHeaderCell())
                'tcl(16).Text = "P1" & "<BR>" & "Remark"
                'tcl(16).BackColor = Color.Red
                'tcl.Add(New TableHeaderCell())

            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then

                For i = 14 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If

        '=========================================================================
        'ARCTERYX
        '=========================================================================
        If pBuyer = "000053" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Type"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "TAS Product"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Centric Code"
                tcl(4).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Finish"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "YKK Code"
                tcl(6).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Description"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Basic Unit"
                tcl(8).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "basic length"
                tcl(9).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "1'' upcharge"
                tcl(10).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "Remark"

            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(6).Text = Replace(e.Row.Cells(6).Text, Chr(10), "<br>")
                For i = 12 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If

        '=========================================================================
        'SALOMON
        '=========================================================================
        If pBuyer = "000055" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Type"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Plasma"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Centric Code"
                tcl(4).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "Buyer Item Description"
                tcl(5).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "YKK Item Description"
                tcl(6).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Basic Length"
                tcl(7).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "Price"
                tcl(8).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "1'' Upcharge"
                tcl(9).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "Remark"

            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 11 To 26
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If


        '=========================================================================
        'TUMI
        '=========================================================================
        If pBuyer = "TW0020" Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "Select"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "A1" & "<BR>" & "Season"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "B1" & "<BR>" & "Type"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "C1" & "<BR>" & "Series"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "D1" & "<BR>" & "Slider/Chain Name"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "E1" & "<BR>" & "TUMI Description"
                tcl(5).BackColor = Color.Green
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "F1" & "<BR>" & "Item Code"
                tcl(6).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "G1" & "<BR>" & "Item Name"
                tcl(7).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "H1" & "<BR>" & "TUMI Color"
                tcl(8).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "I1" & "<BR>" & "SMS Lead Time"
                tcl(9).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(10).Text = "J1" & "<BR>" & "BULK Lead Time"
                tcl(10).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(11).Text = "K1" & "<BR>" & "Unit"
                tcl(11).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(12).Text = "L1" & "<BR>" & "NT$"
                tcl(12).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(13).Text = "M1" & "<BR>" & "US$"
                tcl(13).BackColor = Color.Red
                tcl.Add(New TableHeaderCell())
                tcl(14).Text = "N1" & "<BR>" & "Remark"

            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then

                For i = 15 To 26
                    e.Row.Cells(i).Visible = False
                Next

            End If
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub


End Class
