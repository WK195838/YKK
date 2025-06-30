Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class PS_MENU
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
    Dim NowDate As String           '現在日期時間VVVVV
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
        Response.Cookies("PGM").Value = "PS_MENU.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID

        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        DFileName.Style("left") = -500 & "px"
        DLISTPRICEFile.Style("left") = -500 & "px"
        DBUYERITEMFile.Style("left") = -500 & "px"
        DBUYERCOLORFile.Style("left") = -500 & "px"
        DREPLACEITEMFile.Style("left") = -500 & "px"
        DREPLACECOLORFile.Style("left") = -500 & "px"
        DSALESPRICEFile.Style("left") = -500 & "px"
        DINTEGRATEPRICEFile.Style("left") = -500 & "px"
        DINQLISTPRICEFile.Style("left") = -500 & "px"
        DISOS2FASFile.Style("left") = -500 & "px"
        DINQBUYERITEMFile.Style("left") = -500 & "px"
        DINQBUYERCOLORFile.Style("left") = -500 & "px"
        DItemValuationFile.Style("left") = -500 & "px"
        DITEMSUITABLEFile.Style("left") = -500 & "px"
        DITEMSUITABLEIRWFile.Style("left") = -500 & "px"

        '動作按鈕設定
        BUPLISTPRICE.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('LISTPRICEExcel')}"
        BUPBUYERITEM.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('BUYERITEMExcel')}"
        BUPBUYERCOLOR.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('BUYERCOLORExcel')}"
        BUPREPLACEITEM.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('REPLACEITEMExcel')}"
        BUPREPLACECOLOR.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('REPLACECOLORExcel')}"
        BSALESPRICE.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('SALESPRICEExcel')}"
        BINTEGRATEPRICE.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('INTEGRATEPRICEExcel')}"
        BITEMSUITABLE.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('ITEMSUITABLEExcel')}"

        BTOOLBOX.Attributes.Add("onclick", "window.open('" + "file://10.245.0.186/Share/SystemSupport/[ISOS]ToolBox" + "','_blank');return false;")

        If InStr("SL032/SL034/MK045/IT003/IT004" & "MK055/", UCase(UserID)) > 0 Then
        Else
            'BPRICEVAL.Style("left") = -500 & "px"
            'BIRW.Style("left") = -500 & "px"
            'BITEMTDF00.Style("left") = -500 & "px"
            'BITEMSUITABLE.Style("left") = -500 & "px"
            'BISIP.Style("left") = -500 & "px"
            'BShopping.Style("left") = -500 & "px"
        End If

        'STOP ACT-FCT
        BACTFCT.Style("left") = -500 & "px"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBuyer.Items.Clear()
        'NO
        Dim ListItem1 As New ListItem
        ListItem1.Text = "直接使用"
        ListItem1.Value = "999999"
        ListItem1.Selected = True
        DBuyer.Items.Add(ListItem1)

        'TNF
        Dim ListItem2 As New ListItem
        ListItem2.Text = "TNF"
        ListItem2.Value = "000021"
        DBuyer.Items.Add(ListItem2)
        'ADIDAS
        Dim ListItem3 As New ListItem
        ListItem3.Text = "ADIDAS"
        ListItem3.Value = "000001"
        DBuyer.Items.Add(ListItem3)
        'NIKE
        Dim ListItem4 As New ListItem
        ListItem4.Text = "NIKE"
        ListItem4.Value = "000013"
        DBuyer.Items.Add(ListItem4)
        'COLUMBIA
        Dim ListItem5 As New ListItem
        ListItem5.Text = "COLUMBIA"
        ListItem5.Value = "000003"
        DBuyer.Items.Add(ListItem5)
        'LULULEMON
        Dim ListItem6 As New ListItem
        ListItem6.Text = "LULULEMON"
        ListItem6.Value = "TW1741"
        DBuyer.Items.Add(ListItem6)
        'UNDER ARMOUR
        Dim ListItem7 As New ListItem
        ListItem7.Text = "UNDER ARMOUR"
        ListItem7.Value = "TW0371"
        DBuyer.Items.Add(ListItem7)
        'HELLY HANSEN
        Dim ListItem8 As New ListItem
        ListItem8.Text = "HELLY HANSEN"
        ListItem8.Value = "000098"
        DBuyer.Items.Add(ListItem8)
        'BURTON
        Dim ListItem9 As New ListItem
        ListItem9.Text = "BURTON"
        ListItem9.Value = "000068"
        DBuyer.Items.Add(ListItem9)
        'PUMA
        Dim ListItemA As New ListItem
        ListItemA.Text = "PUMA"
        ListItemA.Value = "000151"
        DBuyer.Items.Add(ListItemA)
        'HERSCHEL
        Dim ListItemB As New ListItem
        ListItemB.Text = "HERSCHEL"
        ListItemB.Value = "TW5068"
        DBuyer.Items.Add(ListItemB)
        'VERA BRADLEY
        Dim ListItemC As New ListItem
        ListItemC.Text = "VERA BRADLEY"
        ListItemC.Value = "TW0655"
        DBuyer.Items.Add(ListItemC)
        'PATAGONIA
        Dim ListItemD As New ListItem
        ListItemD.Text = "PATAGONIA"
        ListItemD.Value = "000141"
        DBuyer.Items.Add(ListItemD)
        'ARCTERYX
        Dim ListItemE As New ListItem
        ListItemE.Text = "ARCTERYX"
        ListItemE.Value = "000053"
        DBuyer.Items.Add(ListItemE)
        'SALOMON
        Dim ListItemF As New ListItem
        ListItemF.Text = "SALOMON"
        ListItemF.Value = "000055"
        DBuyer.Items.Add(ListItemF)
        'TUMI
        Dim ListItemG As New ListItem
        ListItemG.Text = "TUMI"
        ListItemG.Value = "TW0020"
        DBuyer.Items.Add(ListItemG)

        '
        'ISOS-開發G 陳奕儒, 周禹諠, 吳珮婷, JOY, HOWER (SL032/SL034/MK045/IT003/IT004)
        '
        If InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
            'DBuyer.Items.Add(ListItemG)
        End If
        '
        SetTaskButton("ALL", False)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskButton)
    '**     設定各動作按鈕
    '**     pAction=動作名稱    pShow=顯示或不顯示
    '**
    '*****************************************************************
    Sub SetTaskButton(ByVal pAction As String, ByVal pShow As Boolean)
        Select Case pAction
            Case "ALL"
                BUPLISTPRICE.Visible = pShow
                BUPREPLACEITEM.Visible = pShow
                BUPREPLACECOLOR.Visible = pShow
                '
                BDIFFINF.Visible = pShow
                BITEM.Visible = pShow
                BTOOLBOX.Visible = pShow
                BCOLOR.Visible = pShow

                BISIP.Visible = pShow
                'BShopping.Visible = pShow

                BACTFCT.Visible = pShow
                BPRICEVAL.Visible = pShow
                BIRW.Visible = pShow
                BITEMSUITABLE.Visible = pShow
                BITEMTDF00.Visible = pShow
                BITEMINQSA.Visible = pShow
                BITEMDashboard.Visible = pShow

                BDLLISTPRICE.Visible = pShow
                BSALESPRICE.Visible = pShow
                BINTEGRATEPRICE.Visible = pShow
            Case Else
                BUPLISTPRICE.Visible = True
                BUPREPLACEITEM.Visible = True
                BUPREPLACECOLOR.Visible = True
                '
                BDIFFINF.Visible = True
                BITEM.Visible = True
                BTOOLBOX.Visible = True
                BCOLOR.Visible = True

                BPRICEVAL.Visible = True
                BIRW.Visible = True
                BITEMTDF00.Visible = True
                BITEMINQSA.Visible = True
                BITEMSUITABLE.Visible = True
                BITEMDashboard.Visible = True
                BACTFCT.Visible = True
                BISIP.Visible = True
                BShopping.Visible = True

                BDLLISTPRICE.Visible = True
                BSALESPRICE.Visible = True
                BINTEGRATEPRICE.Visible = True
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(選擇BUYER)
    '**     
    '**
    '*****************************************************************
    Protected Sub BBuyer_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBuyer.Click
        Dim xPGM As String
        '
        If DBuyer.SelectedValue <> "999999" And UserID <> "" Then
            '
            ' -- MKT-FUN ------- 
            '共用ITEM & PRICE
            'DLISTPRICEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "LISTPRICE_" & DBuyer.SelectedValue & "_" & UserID & ".xlsm"
            'Select Case DBuyer.SelectedValue
            '    Case "000021"
            '    Case Else
            '        If Not File.Exists(DLISTPRICEFile.Text) Then
            '            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "LISTPRICE_000021.xlsm", DLISTPRICEFile.Text)
            '        End If
            'End Select
            '
            DBUYERITEMFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_" & DBuyer.SelectedValue & "_" & UserID & ".xlsm"
            Select Case DBuyer.SelectedValue
                Case "000021"
                    If InStr("MK013/MK032/SL044", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000021.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000001"
                    If InStr("MK007/SL047/SL048/MK037/MK031/MK055", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000001.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000013"
                    If InStr("MK018/MK025/SL054/MK030", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000013.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000003"
                    If InStr("MK005/SL047", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000003.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "TW1741"
                    If InStr("MK011/MK053/MK006/MK058", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_TW1741.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "TW0371"
                    If InStr("MK013/MK032/MK034/MK030/MK017", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_TW0371.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000151"
                    If InStr("MK006/MK042", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000151.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "TW5068"
                    If InStr("MK048/MK056/MK026", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_TW5068.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "TW0655"
                    If InStr("MK048/MK056/MK026", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_TW0655.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000141"
                    If InStr("MK057/MK053", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000141.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000053"
                    If InStr("MK046/MK034", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000053.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "000055"
                    If InStr("MK046/MK034", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_000055.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case "TW0020"
                    If InStr("TC011", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERITEM_TW0020.xlsm", DBUYERITEMFile.Text)
                        End If
                    End If
                Case Else
            End Select
            '
            DBUYERCOLORFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_" & DBuyer.SelectedValue & "_" & UserID & ".xlsm"
            Select Case DBuyer.SelectedValue
                Case "000021"
                    If InStr("MK013/MK032/SL044", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000021.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000001"
                    If InStr("MK007/SL047/SL048/MK037/MK031/MK055", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000001.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000013"
                    If InStr("MK018/MK025/SL054/MK030/MK007", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000013.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000003"
                    If InStr("MK005/SL047", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000003.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "TW1741"
                    If InStr("MK011/MK053/MK006/MK058", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_TW1741.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "TW0371"
                    If InStr("MK013/MK032/MK034/MK030/MK017", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_TW0371.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000098"
                    If InStr("MK042/MK053", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000098.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000068"
                    If InStr("MK042/MK054/MK038", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000068.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000151"
                    If InStr("MK006/MK042", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000151.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "TW5068"
                    If InStr("MK048/MK026", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_TW5068.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "TW0655"
                    If InStr("MK048/MK026", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_TW0655.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000141"
                    If InStr("MK057/MK053", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000141.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000053"
                    If InStr("MK046/MK017", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000053.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case "000055"
                    If InStr("MK046/MK017", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DBUYERCOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "BUYERCOLOR_000055.xlsm", DBUYERCOLORFile.Text)
                        End If
                    End If
                Case Else
            End Select
            '
            DREPLACEITEMFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "REPLACEITEM_" & DBuyer.SelectedValue & "_" & UserID & ".xlsm"
            Select Case DBuyer.SelectedValue
                Case "000021"
                    '周書羽, 張凱鈞 / 周禹諠, 吳珮婷, JOY
                    If InStr("MK013/MK032/SL044/MK012", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DREPLACEITEMFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "REPLACEITEM_000021.xlsm", DREPLACEITEMFile.Text)
                        End If
                    End If
                Case Else
            End Select
            '
            DREPLACECOLORFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "REPLACECOLOR_" & DBuyer.SelectedValue & "_" & UserID & ".xlsm"
            Select Case DBuyer.SelectedValue
                Case "000021"
                    'MKT 周書羽, 張凱鈞 / 周禹諠, 吳珮婷, JOY
                    If InStr("MK013/MK032/SL044/MK012", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DREPLACECOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "REPLACECOLOR_000021.xlsm", DREPLACECOLORFile.Text)
                        End If
                    End If
                Case "000001"
                    'MKT 周書羽, 張凱鈞 / 周禹諠, 吳珮婷, JOY
                    If InStr("MK007/SL047/SL048/MK037/MK031/MK055", UCase(UserID)) > 0 Or InStr("SL032/SL034/MK045/IT003/IT004", UCase(UserID)) > 0 Then
                        If Not File.Exists(DREPLACECOLORFile.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "REPLACECOLOR_000001.xlsm", DREPLACECOLORFile.Text)
                        End If
                    End If
                Case Else
            End Select
        End If
        '
        If UserID <> "" Then
            '
            ' -- SALES-FUN ------- 
            DSALESPRICEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "SALESPRICE_" & UserID & ".xlsm"
            If Not File.Exists(DSALESPRICEFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SALESPRICE.xlsm", DSALESPRICEFile.Text)
            End If
            '
            DINTEGRATEPRICEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INTEGRATEPRICE_" & UserID & ".xlsm"
            If Not File.Exists(DINTEGRATEPRICEFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INTEGRATEPRICE.xlsm", DINTEGRATEPRICEFile.Text)
            End If
            '
            DINQLISTPRICEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQLISTPRICE_" & UserID & ".xlsm"
            If Not File.Exists(DINQLISTPRICEFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INQLISTPRICE.xlsm", DINQLISTPRICEFile.Text)
            End If
            '
            DISOS2FASFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "ISOS2FAS_" & UserID & ".xlsm"
            If Not File.Exists(DISOS2FASFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "ISOS2FAS.xlsm", DISOS2FASFile.Text)
            End If
            '
            DINQBUYERITEMFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQBUYERITEM_" & UserID & ".xlsm"
            If Not File.Exists(DINQBUYERITEMFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INQBUYERITEM.xlsm", DINQBUYERITEMFile.Text)
            End If
            '
            DINQBUYERCOLORFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQBUYERCOLOR_" & UserID & ".xlsm"
            If Not File.Exists(DINQBUYERCOLORFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INQBUYERCOLOR.xlsm", DINQBUYERCOLORFile.Text)
            End If
            '
            DItemValuationFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "ItemValuation_" & UserID & ".xlsm"
            If Not File.Exists(DItemValuationFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "ItemValuation.xlsm", DItemValuationFile.Text)
            End If
            '
            DITEMSUITABLEFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "ITEMSUITABLE_" & UserID & ".xlsm"
            If Not File.Exists(DITEMSUITABLEFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "ITEMSUITABLE.xlsm", DITEMSUITABLEFile.Text)
            End If
            '
            DITEMSUITABLEIRWFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "ITEMSUITABLEIRW_" & UserID & ".xlsm"
            If Not File.Exists(DITEMSUITABLEIRWFile.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "ITEMSUITABLEIRW.xlsm", DITEMSUITABLEIRWFile.Text)
            End If
            '
            'ISIP-START
            'ADE REPORT
            xPGM = uCommon.GetAppSetting("DataPrepareFile") + "INQPULLERForISIP_" & UserID & ".xlsm"
            If Not File.Exists(xPGM) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INQPULLERForISIP.xlsm", xPGM)
            End If
            'ADV SALES
            xPGM = uCommon.GetAppSetting("DataPrepareFile") + "INQSALESForISIP_" & UserID & ".xlsm"
            If Not File.Exists(xPGM) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INQSALESForISIP.xlsm", xPGM)
            End If
            'PULLER LIST
            xPGM = uCommon.GetAppSetting("DataPrepareFile") + "INQPULLERLISTForISIP_" & UserID & ".xlsm"
            If Not File.Exists(xPGM) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "INQPULLERLISTForISIP.xlsm", xPGM)
            End If
            'ISIP-END
            '
            SetTaskButton("ALL", True)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON ITEM)
    '**     
    '**
    '*****************************************************************
    Protected Sub BDIFFINF_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDIFFINF.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('PS_INQ_UpdateMark.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ITEM','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON ITEM)
    '**     
    '**
    '*****************************************************************
    Protected Sub BITEM_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BITEM.Click
        Dim Cmd As String
        If DBuyer.SelectedValue <> "TW1741" And DBuyer.SelectedValue <> "000151" Then
            Cmd = "<script>" + _
                        "window.open('PS_INQ_Item01.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ITEM','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
                  "</script>"
        Else
            Cmd = "<script>" + _
                        "window.open('PS_INQ_Item03.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ITEM','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
                  "</script>"
        End If
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON COLOR)
    '**     
    '**
    '*****************************************************************
    Protected Sub BCOLOR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BCOLOR.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('PS_INQ_Color01.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','COLOR','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON TOOLBOX)
    '**     
    '**
    '*****************************************************************
    Protected Sub BTOOLBOX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BTOOLBOX.Click
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON ACT-FCT)
    '**     
    '**
    '*****************************************************************
    Protected Sub BACTFCT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BACTFCT.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('ISOS2FAS_AFMain.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ISOS2FAS','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
              "</script>"

        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON ACT-FCT)
    '**     
    '**
    '*****************************************************************
    Protected Sub BISIP_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BISIP.Click
        Dim Cmd, str As String
        If DBuyer.Text = "999999" Then
            Cmd = "<script>" + _
                        "window.open('http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & UserID & "','ISIP','');" + _
                  "</script>"
        Else
            Cmd = "<script>" + _
                        "window.open('http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "&pSource=ISOS" & "','ISIP','');" + _
                  "</script>"
        End If
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON LISTPRICE)
    '**     
    '**
    '*****************************************************************
    Protected Sub BDLLISTPRICE_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDLLISTPRICE.Click
        Dim Cmd As String
        '
        If DBuyer.SelectedValue <> "TW1741" And DBuyer.SelectedValue <> "000151" Then
            Cmd = "<script>" + _
                        "window.open('PS_INQ_ListPrice01.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ListPrice','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
                  "</script>"
        Else
            If DBuyer.SelectedValue = "TW1741" Then
                Cmd = "<script>" + _
                            "window.open('PS_INQ_ListPrice03.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ListPrice','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
                      "</script>"
            Else
                Cmd = "<script>" + _
                            "window.open('PS_INQ_ListPrice04.aspx?pUserID=" & UserID & "&pBuyer=" & DBuyer.Text & "','ListPrice','status=1,toolbar=0,width=1300,height=600,resizable=yes,scrollbars=yes,top=50, left=50');" + _
                      "</script>"
            End If

        End If
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO 預估價)
    '**     
    '**
    '*****************************************************************
    Protected Sub BPRICEVAL_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPRICEVAL.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/ItemValuationSheet_01.aspx?pFormNo=001151&pFormSno=0&pStep=1&pSeqNo=0&pUserID=" & UserID & "&pApplyID=" & UserID & "','預估價','status=1,toolbar=0,width=800,height=800,resizable=yes,scrollbars=yes,top=0, left=0');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO ITEM申請)
    '**     
    '**
    '*****************************************************************
    Protected Sub BITEMTDF00_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BITEMTDF00.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/IRWNOActivity.aspx?pUserID=" & UserID & "','ItemNoAcctivity','status=1,toolbar=0,width=1250,height=800,resizable=yes,scrollbars=yes,top=0, left=0');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO ITEM申請)
    '**     
    '**
    '*****************************************************************
    Protected Sub BIRW_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BIRW.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/ItemRegisterSheet_03.aspx?pFormNo=001151&pFormSno=0&pStep=1&pSeqNo=0&pUserID=" & UserID & "&pApplyID=" & UserID & "','ITEM申請','menubar=yes,width=1200,height=2000,status=1,toolbar=1,resizable=yes,scrollbars=yes,top=0, left=0');" + _
              "</script>"
        Response.Write(Cmd)

        'Cmd = "<script>" + _
        '            "window.open('http://10.245.1.6/IRW/ItemRegisterSheet_03.aspx?pFormNo=001151&pFormSno=0&pStep=1&pSeqNo=0&pUserID=" & UserID & "&pApplyID=" & UserID & "','ITEM申請','status=1,toolbar=0,width=800,height=1200,resizable=yes,scrollbars=yes,top=0, left=0');" + _
        '      "</script>"
        'Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO ITEM-查詢SA)
    '**     
    '**
    '*****************************************************************
    Protected Sub BITEMINQSA_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BITEMINQSA.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/IRWInqItemSA01.aspx?pUserID=" & UserID & "','ITEMINQSA','menubar=yes,width=1200,height=2000,status=1,toolbar=1,resizable=yes,scrollbars=yes,top=0, left=0');" + _
              "</script>"
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO ITEM Dashboard)
    '**     
    '**
    '*****************************************************************
    Protected Sub BITEMDashboard_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BITEMDashboard.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/IRWNoticeMAINPage.aspx?pUserID=" & UserID & "','ITEMINQSA','menubar=yes,width=1200,height=2000,status=1,toolbar=1,resizable=yes,scrollbars=yes,top=0, left=0');" + _
              "</script>"
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO ITEM-查詢模具報廢)
    '**     
    '**
    '*****************************************************************
    Protected Sub BITEMINQMOLD_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BITEMINQMOLD.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/IRWInqMoldCancel.aspx?pUserID=" & UserID & "','ITEMINQMOLD','menubar=yes,width=1200,height=2000,status=1,toolbar=1,resizable=yes,scrollbars=yes,top=0, left=0');" + _
              "</script>"
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO EDX登錄單)
    '**     
    '**
    '*****************************************************************
    Protected Sub BEDXRegister_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BEDXRegister.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('https://forms.office.com/r/MDyLjRE2V2');" + _
              "</script>"
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO Slider Image Portal)
    '**     
    '**
    '*****************************************************************
    Protected Sub BSliderImage3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSliderImage3.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.0.3/ISOS/SliderImageMain.aspx?pUserID=" & UserID & "');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO IT Master)
    '**     
    '**
    '*****************************************************************
    Protected Sub BITMaster_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BITMaster.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/WORKFLOWSUB/INQ_ITMasterStatus.aspx?pUserID=" & UserID & "');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO Shopping System)
    '**     
    '**
    '*****************************************************************
    Protected Sub BShopping_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BShopping.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSSP/StockProgress_01.aspx?pUserID=" & UserID & "');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
End Class
