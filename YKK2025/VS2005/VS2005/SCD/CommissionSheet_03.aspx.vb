Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class CommissionSheet_03
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wNo As String               '委託No
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim NowDateTime As String       '現在日期時間
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        '
        If Not IsPostBack Then                  'PostBack
            SetLinkFile()                       '設定連結檔
            ShowFormData()                      '顯示表單資料
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
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "CommissionSheet_03.aspx"                                   '程式名
        Response.Cookies("NO").Value = Request.QueryString("pNo")                                   '委託No
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
        wNo = Request.QueryString("pNo")                                                            '委託No
        Dim SQL As String
        SQL = "Select * From F_CommissionSheet "
        SQL &= " Where No =  '" & wNo & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then
            wFormNo = dtCommissionSheet.Rows(0).Item("FormNo")
            wFormSno = dtCommissionSheet.Rows(0).Item("FormSno")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetLinkFile)
    '**     設定連結檔
    '**
    '*****************************************************************
    Sub SetLinkFile()
        '-----------------------------------------------------------------
        '-- 開發委託單
        '-----------------------------------------------------------------
        LMAPREFFILE.Visible = False
        LMAPFILE.Visible = False
        LFORTYPEFILE.Visible = False
        LQCCHECKFILE.Visible = False
        LCONTACTFILE.Visible = False
        LSAMPLECONFIRMFILE.Visible = False
        LMANUFAUTHORITYFILE.Visible = False
        LQCFILE1.Visible = False
        LQCFILE2.Visible = False
        LQCFILE3.Visible = False
        LQCFILE4.Visible = False
        LQCFILE5.Visible = False
        LQCFILE6.Visible = False
        '-----------------------------------------------------------------
        '-- 製造委託
        '-----------------------------------------------------------------
        LHINTFILE.Visible = False
        '-----------------------------------------------------------------
        '-- 開發見本
        '-----------------------------------------------------------------
        LSAMPLEFILE.Visible = False
        L3QCFILE1.Visible = False
        L3QCFILE2.Visible = False
        L3QCFILE3.Visible = False
        L3QCFILE4.Visible = False
        L3QCFILE5.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("COMMISSIONFilePath")
        Dim SQL As String
        SQL = "Select * From F_CommissionSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then
            '-----------------------------------------------------------------
            '-- 開發委託單
            '-----------------------------------------------------------------
            '----基本欄位設定-------------------------------------------------    
            DNO.Text = dtCommissionSheet.Rows(0).Item("NO")                             'NO
            If dtCommissionSheet.Rows(0).Item("REFNO") <> "" And _
               (dtCommissionSheet.Rows(0).Item("NeedSample") = 1 Or dtCommissionSheet.Rows(0).Item("NeedItemRegister") = 1) Then       '參考NO
                LREFNO.Text = dtCommissionSheet.Rows(0).Item("REFNO")
                LREFNO.NavigateUrl = uCommon.GetAppSetting("Http") & "CommissionSheet_03.aspx?pNo=" + dtCommissionSheet.Rows(0).Item("REFNO")
                LREFNO.Visible = True
            End If
            DAPPDATE.Text = dtCommissionSheet.Rows(0).Item("APPDATE")                   '申請日
            DAPPDEPT.Text = dtCommissionSheet.Rows(0).Item("APPDEPT")                   '部門
            DAPPPER.Text = dtCommissionSheet.Rows(0).Item("APPPER")                     '職稱
            SetFieldData(0, "APPBUYER", dtCommissionSheet.Rows(0).Item("APPBUYER"))    'BUYER
            DSellVendor.Text = dtCommissionSheet.Rows(0).Item("SELLVENDOR")             '委託廠商
            DESYQTY.Text = dtCommissionSheet.Rows(0).Item("ESYQTY")                     '預估量
            DEXPDEL.Text = dtCommissionSheet.Rows(0).Item("EXPDEL")                     '希望交期
            DCUSTITEM.Text = dtCommissionSheet.Rows(0).Item("CUSTITEM")                 '客戶ITEM
            SetFieldData(0, "USAGE", dtCommissionSheet.Rows(0).Item("USAGE"))          '用途
            DORNO.Text = dtCommissionSheet.Rows(0).Item("ORNO")                         'OR-NO
            SetFieldData(0, "NEEDMAP", dtCommissionSheet.Rows(0).Item("NEEDMAP"))      '需圖面
            '草圖
            If dtCommissionSheet.Rows(0).Item("MAPREFFILE") <> "" Then
                LMAPREFFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MAPREFFILE")
                LMAPREFFILE.Visible = True
            End If
            '----樣品欄位設定-------------------------------------------------    
            If dtCommissionSheet.Rows(0).Item("NeedSample") = 1 Then                                        '需樣品
                DNeedSample.Checked = True
            Else
                DNeedSample.Checked = False
                DNeedSample.Style("top") = -100 & "px"
            End If
            If dtCommissionSheet.Rows(0).Item("NeedItemRegister") = 1 Then                                  '需登錄
                DNeedItemRegister.Checked = True
            Else
                DNeedItemRegister.Checked = False
                DNeedItemRegister.Style("top") = -100 & "px"
            End If
            SetFieldData(0, "PRO", dtCommissionSheet.Rows(0).Item("PRO"))               '製品區分
            DOPPART.Text = dtCommissionSheet.Rows(0).Item("OPPART")                     '開具(色)
            DPLEN.Text = dtCommissionSheet.Rows(0).Item("PLEN")                         '長度(企)
            SetFieldData(0, "PLENUN", dtCommissionSheet.Rows(0).Item("PLENUN"))      '長度單位(企)
            DPQTY.Text = dtCommissionSheet.Rows(0).Item("PQTY")                         '數量(企)
            SetFieldData(0, "PQTYUN", dtCommissionSheet.Rows(0).Item("PQTYUN"))      '數量單位(企)
            DEALEN.Text = dtCommissionSheet.Rows(0).Item("EALEN")                        '長度(EA)
            SetFieldData(0, "EALENUN", dtCommissionSheet.Rows(0).Item("EALENUN")) '長度單位(EA)
            DEAQTY.Text = dtCommissionSheet.Rows(0).Item("EAQTY")                        '數量(EA)
            SetFieldData(0, "EAQTYUN", dtCommissionSheet.Rows(0).Item("EAQTYUN")) '數量單位(EA)
            DUPSLI.Text = dtCommissionSheet.Rows(0).Item("UPSLI")                       '拉頭(上)
            DLOSLI.Text = dtCommissionSheet.Rows(0).Item("LOSLI")                       '拉頭(下)
            DUPFIN.Text = dtCommissionSheet.Rows(0).Item("UPFIN")                       '表面處理(上)
            DLOFIN.Text = dtCommissionSheet.Rows(0).Item("LOFIN")                       '表面處理(下)
            DUPSTK.Text = dtCommissionSheet.Rows(0).Item("UPSTK")                       '上止種類
            DLOSTK.Text = dtCommissionSheet.Rows(0).Item("LOSTK")                       '下止種類
            DSPSPEC.Text = dtCommissionSheet.Rows(0).Item("SPSPEC")                     '特殊規格
            '----開發規格欄位設定-------------------------------------------------    
            SetFieldData(0, "SIZENO", dtCommissionSheet.Rows(0).Item("SIZENO"))     '型別
            SetFieldData(0, "ITEM", dtCommissionSheet.Rows(0).Item("ITEM"))         '鏈條型式
            SetFieldData(0, "TATYPE", dtCommissionSheet.Rows(0).Item("TATYPE"))     '布帶
            DTAWIDTH.Text = dtCommissionSheet.Rows(0).Item("TAWIDTH")                   '布帶寬度
            SetFieldData(0, "ECOLSEL", dtCommissionSheet.Rows(0).Item("ECOLSEL"))    '鏈齒顏色-SEL
            DECOL.Text = dtCommissionSheet.Rows(0).Item("ECOL")                         '鏈齒顏色
            SetFieldData(0, "CCOLSEL", dtCommissionSheet.Rows(0).Item("CCOLSEL"))   '丸紐-SEL
            DCCOL.Text = dtCommissionSheet.Rows(0).Item("CCOL")                         '丸紐
            SetFieldData(0, "TACOL", dtCommissionSheet.Rows(0).Item("TACOL"))       '布帶-色番(同)
            DTACOLNO.Text = dtCommissionSheet.Rows(0).Item("TACOLNO")                   '布帶-色番(同)
            DTAYCOLNO.Text = dtCommissionSheet.Rows(0).Item("TAYCOLNO")                 '布帶-YKK(同)
            SetFieldData(0, "TALCOL", dtCommissionSheet.Rows(0).Item("TALCOL"))     '布帶-色番(左)
            DTALCOLNO.Text = dtCommissionSheet.Rows(0).Item("TALCOLNO")                 '布帶-色番(左)
            DTALYCOLNO.Text = dtCommissionSheet.Rows(0).Item("TALYCOLNO")               '布帶-YKK(左)
            SetFieldData(0, "TARCOL", dtCommissionSheet.Rows(0).Item("TARCOL"))     '布帶-色番(右)
            DTARCOLNO.Text = dtCommissionSheet.Rows(0).Item("TARCOLNO")                 '布帶-色番(右)
            DTARYCOLNO.Text = dtCommissionSheet.Rows(0).Item("TARYCOLNO")               '布帶-YKK(右)
            SetFieldData(0, "THUPCOL", dtCommissionSheet.Rows(0).Item("THUPCOL"))   '縫上-色番(同)
            DTHUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THUPCOLNO")               '縫上-色番(同)
            DTHUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THUPYCOLNO")             '縫上-YKK(同)
            SetFieldData(0, "THLUPCOL", dtCommissionSheet.Rows(0).Item("THLUPCOL")) '縫上-色番(左)
            DTHLUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPCOLNO")             '縫上-色番(左)
            DTHLUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPYCOLNO")           '縫上-YKK(左)

            SetFieldData(0, "THLRUPCOL", dtCommissionSheet.Rows(0).Item("THLRUPCOL")) '縫上-色番(左右)
            DTHLRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPCOLNO")             '縫上-色番(左右)
            DTHLRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPYCOLNO")           '縫上-YKK(左右)


            SetFieldData(0, "THRUPCOL", dtCommissionSheet.Rows(0).Item("THRUPCOL")) '縫上-色番(右)
            DTHRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPCOLNO")             '縫上-色番(右)
            DTHRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPYCOLNO")           '縫上-YKK(右)

            SetFieldData(0, "THRRUPCOL", dtCommissionSheet.Rows(0).Item("THRRUPCOL")) '縫上-色番(右右)
            DTHRRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRUPCOLNO")             '縫上-色番(右右)
            DTHRRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRUPYCOLNO")           '縫上-YKK(右右)


            SetFieldData(0, "THLOCOL", dtCommissionSheet.Rows(0).Item("THLOCOL"))   '縫下-色番(同)
            DTHLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOCOLNO")               '縫下-色番(同)
            DTHLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOYCOLNO")             '縫下-YKK(同)
            SetFieldData(0, "THLLOCOL", dtCommissionSheet.Rows(0).Item("THLLOCOL")) '縫下-色番(左)
            DTHLLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOCOLNO")             '縫下-色番(左)
            DTHLLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOYCOLNO")           '縫下-YKK(左)

            SetFieldData(0, "THLRLOCOL", dtCommissionSheet.Rows(0).Item("THLRLOCOL")) '縫下-色番(左右)
            DTHLRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRLOCOLNO")             '縫下-色番(左右)
            DTHLRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRLOYCOLNO")           '縫下-YKK(左右)


            SetFieldData(0, "THRLOCOL", dtCommissionSheet.Rows(0).Item("THRLOCOL")) '縫下-色番(右)
            DTHRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRLOCOLNO")             '縫下-色番(右)
            DTHRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRLOYCOLNO")           '縫下-YKK(右)


            SetFieldData(0, "THRRLOCOL", dtCommissionSheet.Rows(0).Item("THRRLOCOL")) '縫下-色番(右右)
            DTHRRLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRLOCOLNO")             '縫下-色番(右右)
            DTHRRLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRLOYCOLNO")           '縫下-YKK(右右)



            DXMLEN.Text = dtCommissionSheet.Rows(0).Item("XMLEN")                       'X-尺寸
            SetFieldData(0, "XMCOL", dtCommissionSheet.Rows(0).Item("XMCOL"))       'X-色番
            DXMCOLNO.Text = dtCommissionSheet.Rows(0).Item("XMCOLNO")                   'X-色番
            DXMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("XMYCOLNO")                 'X-YKK
            DAMLEN.Text = dtCommissionSheet.Rows(0).Item("AMLEN")                       'A-尺寸
            SetFieldData(0, "AMCOL", dtCommissionSheet.Rows(0).Item("AMCOL"))       'A-色番
            DAMCOLNO.Text = dtCommissionSheet.Rows(0).Item("AMCOLNO")                   'A-色番
            DAMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("AMYCOLNO")                 'A-YKK
            DBMLEN.Text = dtCommissionSheet.Rows(0).Item("BMLEN")                       'B-尺寸
            SetFieldData(0, "BMCOL", dtCommissionSheet.Rows(0).Item("BMCOL"))       'B-色番
            DBMCOLNO.Text = dtCommissionSheet.Rows(0).Item("BMCOLNO")                   'B-色番
            DBMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("BMYCOLNO")                 'B-YKK
            DCMLEN.Text = dtCommissionSheet.Rows(0).Item("CMLEN")                       'C-尺寸
            SetFieldData(0, "CMCOL", dtCommissionSheet.Rows(0).Item("CMCOL"))       'C-色番
            DCMCOLNO.Text = dtCommissionSheet.Rows(0).Item("CMCOLNO")                   'C-色番
            DCMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("CMYCOLNO")                 'C-YKK
            DDMLEN.Text = dtCommissionSheet.Rows(0).Item("DMLEN")                       'D-尺寸
            SetFieldData(0, "DMCOL", dtCommissionSheet.Rows(0).Item("DMCOL"))       'D-色番
            DDMCOLNO.Text = dtCommissionSheet.Rows(0).Item("DMCOLNO")                   'D-色番
            DDMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("DMYCOLNO")                 'D-YKK
            DEMLEN.Text = dtCommissionSheet.Rows(0).Item("EMLEN")                       'E-尺寸
            SetFieldData(0, "EMCOL", dtCommissionSheet.Rows(0).Item("EMCOL"))       'E-色番
            DEMCOLNO.Text = dtCommissionSheet.Rows(0).Item("EMCOLNO")                   'E-色番
            DEMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("EMYCOLNO")                 'E-YKK
            DFMLEN.Text = dtCommissionSheet.Rows(0).Item("FMLEN")                       'F-尺寸
            SetFieldData(0, "FMCOL", dtCommissionSheet.Rows(0).Item("FMCOL"))       'F-色番
            DFMCOLNO.Text = dtCommissionSheet.Rows(0).Item("FMCOLNO")                   'F-色番
            DFMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("FMYCOLNO")                 'F-YKK
            DGMLEN.Text = dtCommissionSheet.Rows(0).Item("GMLEN")                       'G-尺寸
            SetFieldData(0, "GMCOL", dtCommissionSheet.Rows(0).Item("GMCOL"))       'G-色番
            DGMCOLNO.Text = dtCommissionSheet.Rows(0).Item("GMCOLNO")                   'G-色番
            DGMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("GMYCOLNO")                 'G-YKK
            DHMLEN.Text = dtCommissionSheet.Rows(0).Item("HMLEN")                       'H-尺寸
            SetFieldData(0, "HMCOL", dtCommissionSheet.Rows(0).Item("HMCOL"))       'H-色番
            DHMCOLNO.Text = dtCommissionSheet.Rows(0).Item("HMCOLNO")                   'H-色番
            DHMYCOLNO.Text = dtCommissionSheet.Rows(0).Item("HMYCOLNO")                 'H-YKK
            DLYLEN.Text = dtCommissionSheet.Rows(0).Item("LYLEN")                       '緯紗-尺寸
            SetFieldData(0, "LYCOL", dtCommissionSheet.Rows(0).Item("LYCOL"))       '緯紗-色番
            DLYCOLNO.Text = dtCommissionSheet.Rows(0).Item("LYCOLNO")                   '緯紗-色番
            DLYYCOLNO.Text = dtCommissionSheet.Rows(0).Item("LYYCOLNO")                 '緯紗-YKK
            DOTCON.Text = dtCommissionSheet.Rows(0).Item("OTCON")                       '其他條件
            '----圖面欄位設定-------------------------------------------------    
            DMAPNO.Text = dtCommissionSheet.Rows(0).Item("MAPNO")                       '圖號
            SetFieldData(0, "MAKEMAP", dtCommissionSheet.Rows(0).Item("MAKEMAP"))       '製圖者
            SetFieldData(0, "LEVEL", dtCommissionSheet.Rows(0).Item("LEVEL"))           '難易度
            '圖檔
            If dtCommissionSheet.Rows(0).Item("MAPFILE") <> "" Then
                LMAPFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MAPFILE")
                LMAPFILE.Visible = True
            End If
            '----其他附件-------------------------------------------------    
            '適用型別檔
            If dtCommissionSheet.Rows(0).Item("FORTYPEFILE") <> "" Then
                LFORTYPEFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORTYPEFILE")
                LFORTYPEFILE.Visible = True
            End If
            '品質檢測項目檔
            If dtCommissionSheet.Rows(0).Item("QCCHECKFILE") <> "" Then
                LQCCHECKFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCCHECKFILE")
                LQCCHECKFILE.Visible = True
            End If

            'QC檢測文件
            If dtCommissionSheet.Rows(0).Item("QCFILE1") <> "" Then
                LQCFILE1.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE1")
                LQCFILE1.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE2") <> "" Then
                LQCFILE2.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE2")
                LQCFILE2.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE3") <> "" Then
                LQCFILE3.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE3")
                LQCFILE3.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE4") <> "" Then
                LQCFILE4.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE4")
                LQCFILE4.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE5") <> "" Then
                LQCFILE5.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE5")
                LQCFILE5.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE6") <> "" Then
                LQCFILE6.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE6")
                LQCFILE6.Visible = True
            End If

            '客戶切結書
            If dtCommissionSheet.Rows(0).Item("CONTACTFILE") <> "" Then
                LCONTACTFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("CONTACTFILE")
                LCONTACTFILE.Visible = True
            End If
            '樣品確認書
            If dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE") <> "" Then
                LSAMPLECONFIRMFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE")
                LSAMPLECONFIRMFILE.Visible = True
            End If
            '製造授權書
            If dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE") <> "" Then
                LMANUFAUTHORITYFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE")
                LMANUFAUTHORITYFILE.Visible = True
            End If

            DMANUOUTPRICE.Text = dtCommissionSheet.Rows(0).Item("MANUOUTPRICE")     '外注加工費

            '報價單
            If dtCommissionSheet.Rows(0).Item("FORCASTFILE") <> "" Then
                LFORCASTFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORCASTFILE")
                LFORCASTFILE.Visible = True
            End If
            '-----------------------------------------------------------------
            '-- 製造委託
            '-----------------------------------------------------------------
            SQL = "Select * From FS_ManufactureSheet "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            Dim dtManufactureSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dtManufactureSheet.Rows.Count > 0 Then
                '----基本欄位設定-------------------------------------------------    
                DDEVTITLE.Text = dtManufactureSheet.Rows(0).Item("DEVTITLE")            '開發主題
                DDEVNO.Text = dtManufactureSheet.Rows(0).Item("DEVNO")                  '開發NO.
                DCODENO.Text = dtManufactureSheet.Rows(0).Item("CODENO")                'CODE NO.
                DISSDATE.Text = dtManufactureSheet.Rows(0).Item("ISSDATE")              '發行日
                DDEVPER1.Text = dtManufactureSheet.Rows(0).Item("DEVPER1")              '開發擔當
                '----開發內容-------------------------------------------------    
                '示意圖檔
                If dtManufactureSheet.Rows(0).Item("HINTFILE") <> "" Then
                    LHINTFILE.ImageUrl = Path & dtManufactureSheet.Rows(0).Item("HINTFILE")
                    LHINTFILE.Visible = True
                End If
                DUPSTK1.Text = dtManufactureSheet.Rows(0).Item("UPSTK")                  '上止
                DLOSTK1.Text = dtManufactureSheet.Rows(0).Item("LOSTK")                  '下止
                DOPPART1.Text = dtManufactureSheet.Rows(0).Item("OPPART")                '開具(色)
                DTASPEC.Text = dtManufactureSheet.Rows(0).Item("TASPEC")                '布帶
                DECOL1.Text = dtManufactureSheet.Rows(0).Item("ECOL")                    '鏈齒顏色
                DCCOL1.Text = dtManufactureSheet.Rows(0).Item("CCOL")                    '丸紐
                DTHSPEC.Text = dtManufactureSheet.Rows(0).Item("THSPEC")                '縫工線
                DPLEN1.Text = dtManufactureSheet.Rows(0).Item("PLEN")                    '長度(企)
                DPQTY1.Text = dtManufactureSheet.Rows(0).Item("PQTY")                    '數量(企)
                DEALEN1.Text = dtManufactureSheet.Rows(0).Item("EALEN")                  '長度(EA)
                DEAQTY1.Text = dtManufactureSheet.Rows(0).Item("EAQTY")                  '數量(EA)
                '----工程-------------------------------------------------    
                DMANUFTYPE.Text = dtManufactureSheet.Rows(0).Item("MANUFTYPE")          '內外製
                HOP1.Text = dtManufactureSheet.Rows(0).Item("OP1")                      'OP1-工程
                HOP1.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP1.Text)
                DOP1PER.Text = dtManufactureSheet.Rows(0).Item("OP1PER")                'OP1-擔當
                DOP1BTIME.Text = dtManufactureSheet.Rows(0).Item("OP1BTIME")            'OP1-預定納期
                DOP1BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP1BHOURS")          'OP1-預定時數
                DOP1ATIME.Text = dtManufactureSheet.Rows(0).Item("OP1ATIME")            'OP1-實際納期
                DOP1AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP1AHOURS")          'OP1-實際時數
                DOP1CON.Text = dtManufactureSheet.Rows(0).Item("OP1CON")                'OP1-作業內容
                SetFieldData(0, "OP1DELAYC1", dtManufactureSheet.Rows(0).Item("OP1DELAYC1"))       'OP1-遲納原因-1
                SetFieldData(0, "OP1DELAYC2", dtManufactureSheet.Rows(0).Item("OP1DELAYC2"))       'OP1-遲納原因-2
                DOP1REM.Text = dtManufactureSheet.Rows(0).Item("OP1REM")                'OP1-遲納原因

                HOP2.Text = dtManufactureSheet.Rows(0).Item("OP2")                      'OP2-工程
                HOP2.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP2.Text)
                DOP2PER.Text = dtManufactureSheet.Rows(0).Item("OP2PER")                'OP2-擔當
                DOP2BTIME.Text = dtManufactureSheet.Rows(0).Item("OP2BTIME")            'OP2-預定納期
                DOP2BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP2BHOURS")          'OP2-預定時數
                DOP2ATIME.Text = dtManufactureSheet.Rows(0).Item("OP2ATIME")            'OP2-實際納期
                DOP2AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP2AHOURS")          'OP2-實際時數
                DOP2CON.Text = dtManufactureSheet.Rows(0).Item("OP2CON")                'OP2-作業內容
                SetFieldData(0, "OP2DELAYC1", dtManufactureSheet.Rows(0).Item("OP2DELAYC1"))       'OP2-遲納原因-1
                SetFieldData(0, "OP2DELAYC2", dtManufactureSheet.Rows(0).Item("OP2DELAYC2"))       'OP2-遲納原因-2
                DOP2REM.Text = dtManufactureSheet.Rows(0).Item("OP2REM")                'OP2-遲納原因

                HOP3.Text = dtManufactureSheet.Rows(0).Item("OP3")                      'OP3-工程
                HOP3.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP3.Text)
                DOP3PER.Text = dtManufactureSheet.Rows(0).Item("OP3PER")                'OP3-擔當
                DOP3BTIME.Text = dtManufactureSheet.Rows(0).Item("OP3BTIME")            'OP3-預定納期
                DOP3BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP3BHOURS")          'OP3-預定時數
                DOP3ATIME.Text = dtManufactureSheet.Rows(0).Item("OP3ATIME")            'OP3-實際納期
                DOP3AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP3AHOURS")          'OP3-實際時數
                DOP3CON.Text = dtManufactureSheet.Rows(0).Item("OP3CON")                'OP3-作業內容
                SetFieldData(0, "OP3DELAYC1", dtManufactureSheet.Rows(0).Item("OP3DELAYC1"))       'OP3-遲納原因-1
                SetFieldData(0, "OP3DELAYC2", dtManufactureSheet.Rows(0).Item("OP3DELAYC2"))       'OP3-遲納原因-2
                DOP3REM.Text = dtManufactureSheet.Rows(0).Item("OP3REM")                'OP3-遲納原因

                HOP4.Text = dtManufactureSheet.Rows(0).Item("OP4")                      'OP4-工程
                HOP4.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP4.Text)
                DOP4PER.Text = dtManufactureSheet.Rows(0).Item("OP4PER")                'OP4-擔當
                DOP4BTIME.Text = dtManufactureSheet.Rows(0).Item("OP4BTIME")            'OP4-預定納期
                DOP4BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP4BHOURS")          'OP4-預定時數
                DOP4ATIME.Text = dtManufactureSheet.Rows(0).Item("OP4ATIME")            'OP4-實際納期
                DOP4AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP4AHOURS")          'OP4-實際時數
                DOP4CON.Text = dtManufactureSheet.Rows(0).Item("OP4CON")                'OP4-作業內容
                SetFieldData(0, "OP4DELAYC1", dtManufactureSheet.Rows(0).Item("OP4DELAYC1"))       'OP4-遲納原因-1
                SetFieldData(0, "OP4DELAYC2", dtManufactureSheet.Rows(0).Item("OP4DELAYC2"))       'OP4-遲納原因-2
                DOP4REM.Text = dtManufactureSheet.Rows(0).Item("OP4REM")                'OP4-遲納原因

                HOP5.Text = dtManufactureSheet.Rows(0).Item("OP5")                      'OP5-工程
                HOP5.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP5.Text)
                DOP5PER.Text = dtManufactureSheet.Rows(0).Item("OP5PER")                'OP5-擔當
                DOP5BTIME.Text = dtManufactureSheet.Rows(0).Item("OP5BTIME")            'OP5-預定納期
                DOP5BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP5BHOURS")          'OP5-預定時數
                DOP5ATIME.Text = dtManufactureSheet.Rows(0).Item("OP5ATIME")            'OP5-實際納期
                DOP5AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP5AHOURS")          'OP5-實際時數
                DOP5CON.Text = dtManufactureSheet.Rows(0).Item("OP5CON")                'OP5-作業內容
                SetFieldData(0, "OP5DELAYC1", dtManufactureSheet.Rows(0).Item("OP5DELAYC1"))       'OP5-遲納原因-1
                SetFieldData(0, "OP5DELAYC2", dtManufactureSheet.Rows(0).Item("OP5DELAYC2"))       'OP5-遲納原因-2
                DOP5REM.Text = dtManufactureSheet.Rows(0).Item("OP5REM")                'OP5-遲納原因

                HOP6.Text = dtManufactureSheet.Rows(0).Item("OP6")                      'OP6-工程
                HOP6.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP6.Text)
                DOP6PER.Text = dtManufactureSheet.Rows(0).Item("OP6PER")                'OP6-擔當
                DOP6BTIME.Text = dtManufactureSheet.Rows(0).Item("OP6BTIME")            'OP6-預定納期
                DOP6BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP6BHOURS")          'OP6-預定時數
                DOP6ATIME.Text = dtManufactureSheet.Rows(0).Item("OP6ATIME")            'OP6-實際納期
                DOP6AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP6AHOURS")          'OP6-實際時數
                DOP6CON.Text = dtManufactureSheet.Rows(0).Item("OP6CON")                'OP6-作業內容
                SetFieldData(0, "OP6DELAYC1", dtManufactureSheet.Rows(0).Item("OP6DELAYC1"))       'OP6-遲納原因-1
                SetFieldData(0, "OP6DELAYC2", dtManufactureSheet.Rows(0).Item("OP6DELAYC2"))       'OP6-遲納原因-2
                DOP6REM.Text = dtManufactureSheet.Rows(0).Item("OP6REM")                'OP6-遲納原因

                HOP7.Text = dtManufactureSheet.Rows(0).Item("OP7")                      'OP7-工程
                HOP7.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP7.Text)
                DOP7PER.Text = dtManufactureSheet.Rows(0).Item("OP7PER")                'OP7-擔當
                DOP7BTIME.Text = dtManufactureSheet.Rows(0).Item("OP7BTIME")            'OP7-預定納期
                DOP7BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP7BHOURS")          'OP7-預定時數
                DOP7ATIME.Text = dtManufactureSheet.Rows(0).Item("OP7ATIME")            'OP7-實際納期
                DOP7AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP7AHOURS")          'OP7-實際時數
                DOP7CON.Text = dtManufactureSheet.Rows(0).Item("OP7CON")                'OP7-作業內容
                SetFieldData(0, "OP7DELAYC1", dtManufactureSheet.Rows(0).Item("OP7DELAYC1"))       'OP7-遲納原因-1
                SetFieldData(0, "OP7DELAYC2", dtManufactureSheet.Rows(0).Item("OP7DELAYC2"))       'OP7-遲納原因-2
                DOP7REM.Text = dtManufactureSheet.Rows(0).Item("OP7REM")                'OP7-遲納原因

                HOP8.Text = dtManufactureSheet.Rows(0).Item("OP8")                      'OP8-工程
                HOP8.NavigateUrl = "ManufactureHistory.aspx?&pFormNo=" & wFormNo & "&pFormSno=" & CStr(wFormSno) & "&pOPCode=" & fpObj.GetOPCode("CODE-" & HOP8.Text)
                DOP8PER.Text = dtManufactureSheet.Rows(0).Item("OP8PER")                'OP8-擔當
                DOP8BTIME.Text = dtManufactureSheet.Rows(0).Item("OP8BTIME")            'OP8-預定納期
                DOP8BHOURS.Text = dtManufactureSheet.Rows(0).Item("OP8BHOURS")          'OP8-預定時數
                DOP8ATIME.Text = dtManufactureSheet.Rows(0).Item("OP8ATIME")            'OP8-實際納期
                DOP8AHOURS.Text = dtManufactureSheet.Rows(0).Item("OP8AHOURS")          'OP8-實際時數
                DOP8CON.Text = dtManufactureSheet.Rows(0).Item("OP8CON")                'OP8-作業內容
                SetFieldData(0, "OP8DELAYC1", dtManufactureSheet.Rows(0).Item("OP8DELAYC1"))       'OP8-遲納原因-1
                SetFieldData(0, "OP8DELAYC2", dtManufactureSheet.Rows(0).Item("OP8DELAYC2"))       'OP8-遲納原因-2
                DOP8REM.Text = dtManufactureSheet.Rows(0).Item("OP8REM")                'OP8-遲納原因
            End If
            '-----------------------------------------------------------------
            '-- 開發見本
            '-----------------------------------------------------------------
            SQL = "Select * From FS_SampleSheet "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dtSampleSheet.Rows.Count > 0 Then
                '----基本欄位設定-------------------------------------------------
                D3APPBUYER.Text = dtSampleSheet.Rows(0).Item("AppBuyer")                 'Customer
                D3DATE.Text = dtSampleSheet.Rows(0).Item("Date")                         '發行日
                D3SIZENO.Text = dtSampleSheet.Rows(0).Item("SizeNo")                     'Size
                D3ITEM.Text = dtSampleSheet.Rows(0).Item("Item")                         'Item
                D3CODENO.Text = dtSampleSheet.Rows(0).Item("CodeNo")                     'Code No
                '實際樣品
                If dtSampleSheet.Rows(0).Item("SampleFile") <> "" Then
                    LSAMPLEFILE.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile")
                    LSAMPLEFILE.Visible = True
                End If
                '----開發規格-------------------------------------------------
                D3TAWIDTH.Text = dtSampleSheet.Rows(0).Item("TAWidth")                   '布帶寬度
                D3DEVNO.Text = dtSampleSheet.Rows(0).Item("DevNo")                       '開發No
                D3DEVPRD.Text = dtSampleSheet.Rows(0).Item("DevPrd")                     '開發期間
                D3TACOL.Text = dtSampleSheet.Rows(0).Item("TACol")                       '布帶Color
                D3TALINE.Text = dtSampleSheet.Rows(0).Item("TALine")                     '條紋線Color
                D3ECOL.Text = dtSampleSheet.Rows(0).Item("ECol")                         '務齒
                D3CCOL.Text = dtSampleSheet.Rows(0).Item("CCol")                         '丸紐
                D3THCOL.Text = dtSampleSheet.Rows(0).Item("THCol")                       '縫工線
                D3OTHER.Text = dtSampleSheet.Rows(0).Item("Other")                       '其他
                '----QC附檔-------------------------------------------------
                If dtSampleSheet.Rows(0).Item("QCFile1") <> "" Then                     '品測報告1
                    L3QCFILE1.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile1")
                    L3QCFILE1.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile2") <> "" Then                     '品測報告2
                    L3QCFILE2.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile2")
                    L3QCFILE2.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile3") <> "" Then                     '品測報告3
                    L3QCFILE3.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile3")
                    L3QCFILE3.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile4") <> "" Then                     '品測報告4
                    L3QCFILE4.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile4")
                    L3QCFILE4.Visible = True
                End If
                If dtSampleSheet.Rows(0).Item("QCFile5") <> "" Then                     '品測報告5
                    L3QCFILE5.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile5")
                    L3QCFILE5.Visible = True
                End If
                '----Wave's-------------------------------------------------
                D3TNLITEM.Text = dtSampleSheet.Rows(0).Item("TNLItem")                   'TAPE NAT-左
                D3TNRITEM.Text = dtSampleSheet.Rows(0).Item("TNRItem")                   'TAPE NAT-右
                D3TSLITEM.Text = dtSampleSheet.Rows(0).Item("TSLItem")                   'TAPE SET-左
                D3TSRITEM.Text = dtSampleSheet.Rows(0).Item("TSRItem")                   'TAPE SET-右
                D3TDLITEM.Text = dtSampleSheet.Rows(0).Item("TDLItem")                   'TAPE DYED-左
                D3TDRITEM.Text = dtSampleSheet.Rows(0).Item("TDRItem")                   'TAPE DYED-右
                D3CNITEM.Text = dtSampleSheet.Rows(0).Item("CNItem")                     'CHAIN NAT
                D3CSITEM.Text = dtSampleSheet.Rows(0).Item("CSItem")                     'CHAIN SET
                D3CDITEM.Text = dtSampleSheet.Rows(0).Item("CDItem")                     'CHAIN DYED
                D31Other.Text = dtSampleSheet.Rows(0).Item("Other1")                     'Other1
                D32Other.Text = dtSampleSheet.Rows(0).Item("Other2")                     'Other2
                D3O1ITEM.Text = dtSampleSheet.Rows(0).Item("O1Item")                     'O1Item
                D3O2ITEM.Text = dtSampleSheet.Rows(0).Item("O2Item")                     'O2Item
                D3CITEM.Text = dtSampleSheet.Rows(0).Item("CItem")                       'CORD
                '----FLOW-------------------------------------------------
                SetFieldData(0, "WF1", dtSampleSheet.Rows(0).Item("WF1"))          '承認WF1
                SetFieldData(0, "WF2", dtSampleSheet.Rows(0).Item("WF2"))          '承認WF2
                SetFieldData(0, "WF3", dtSampleSheet.Rows(0).Item("WF3"))          '承認WF3
                SetFieldData(0, "WF4", dtSampleSheet.Rows(0).Item("WF4"))          '承認WF4
                SetFieldData(0, "WF5", dtSampleSheet.Rows(0).Item("WF5"))          '承認WF5
                SetFieldData(0, "WF6", dtSampleSheet.Rows(0).Item("WF6"))          '承認WF6
                SetFieldData(0, "WF7", dtSampleSheet.Rows(0).Item("WF7"))          '承認WF7
                SetFieldData(0, "WF3NAME", dtSampleSheet.Rows(0).Item("WF3Name"))  '承認者部門WF3
                SetFieldData(0, "WF4NAME", dtSampleSheet.Rows(0).Item("WF4Name"))  '承認者部門WF4
                SetFieldData(0, "WF5NAME", dtSampleSheet.Rows(0).Item("WF5Name"))  '承認者部門WF5
                SetFieldData(0, "WF6NAME", dtSampleSheet.Rows(0).Item("WF6Name"))  '承認者部門WF6
                SetFieldData(0, "WF7NAME", dtSampleSheet.Rows(0).Item("WF7Name"))  '承認者部門WF7
            End If
            '-----------------------------------------------------------------
            '-- 原單位
            '-----------------------------------------------------------------
            SQL = "Select * From FS_GentaniSheet "
            SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
            Dim dtGentaniSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dtGentaniSheet.Rows.Count > 0 Then
                '
                Dim s As New List(Of String)
                s.Add("TNLITEM")
                s.Add("TNRITEM")
                s.Add("ECOL")
                s.Add("EITEM")
                '
                For Each dc As DataColumn In dtGentaniSheet.Columns
                    Dim l As Object = Me.form1.FindControl("D4" & dc.ColumnName)
                    Dim v As String = uCommon.ReplaceDBnull(dtGentaniSheet.Rows(0)(dc.ColumnName), "")
                    If l IsNot Nothing Then
                        l.Text = v
                    Else
                        If s.Contains(dc.ColumnName) Then
                            l = Me.form1.FindControl("D4" & dc.ColumnName & "1")
                            l.Text = v
                            l = Me.form1.FindControl("D4" & dc.ColumnName & "2")
                            l.Text = v
                        End If
                    End If
                Next
                '
            End If
            '-----------------------------------------------------------------
            '-- 核定履歷
            '-----------------------------------------------------------------
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'Modify-Start by Joy  2012/7/31 新納期對應
            '
            'SQL = "SELECT "
            'SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            'SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            'SQL = SQL + "'預定：[' + BStartTimeDesc + ' ~ ' + BEndTimeDesc + '], ' + "
            'SQL = SQL + "'實際：[' + AStartTimeDesc + ' ~ ' + AEndTimeDesc + '] ' As Description, "
            'SQL = SQL + "URL "
            'SQL = SQL + "FROM V_WaitHandle_01 "
            'SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            'SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            'GridView1.DataSource = uDataBase.GetDataTable(SQL)
            'GridView1.DataBind()
            '
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "'' as Description, BStartTimeDesc, BEndTimeDesc, AStartTimeDesc, AEndTimeDesc, WorkID, "
            SQL = SQL + "'' As ViewOP, '' AS URL, Active "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            SQL = SQL + "  And Active <> '9' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dt_WaitHandle.Rows.Count > 0 Then
                For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
                    ' StepNameDesc(9)
                    dt_WaitHandle.Rows(i)(9) = "[" + dt_WaitHandle.Rows(i)(2).ToString + "]-" + Replace(dt_WaitHandle.Rows(i)(9), "開發委託_", "")
                    '
                    ' Description(12)
                    ' 預：[2012-04-25 16:00 ~ 2012-04-25 17:00](380),
                    ' 實：[2012-04-25 16:00 ~ 2012-04-26 14:10](380)
                    Dim xTime As Integer = 0
                    ' 預
                    xTime = 0
                    If InStr(dt_WaitHandle.Rows(i)(14), "無記錄") > 0 Then
                    Else
                        dt_WaitHandle.Rows(i)(12) = "預：[" + CDate(dt_WaitHandle.Rows(i)(13)).ToString("yyyy/MM/dd HH:mm") + " ~ " + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyy/MM/dd HH:mm")
                        oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i)(13)).ToString("yyyy/MM/dd HH:mm"), _
                                                 CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyy/MM/dd HH:mm"), _
                                                 xTime, _
                                                 oCommon.GetCalendarGroup(dt_WaitHandle.Rows(i)(17)))
                        dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + "  (" + CStr(xTime) + ")]"
                    End If
                    ' 實
                    xTime = 0
                    If InStr(dt_WaitHandle.Rows(i)(16), "無記錄") > 0 Then
                    Else
                        dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + ", " + _
                                                    "實：[" + CDate(dt_WaitHandle.Rows(i)(15)).ToString("yyyy/MM/dd HH:mm") + " ~ " + CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyy/MM/dd HH:mm")
                        oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i)(15)).ToString("yyyy/MM/dd HH:mm"), _
                                                 CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyy/MM/dd HH:mm"), _
                                                 xTime, _
                                                 oCommon.GetCalendarGroup(dt_WaitHandle.Rows(i)(17)))
                        dt_WaitHandle.Rows(i)(12) = dt_WaitHandle.Rows(i)(12) + "  (" + CStr(xTime) + ")]"
                    End If
                    ' ViewOP(18), URL(19)
                    If GetFlowLoading(dt_WaitHandle.Rows(i)(2)) = 1 Then
                        dt_WaitHandle.Rows(i)(18) = "@"
                        If dt_WaitHandle.Rows(i)(20) <> 1 Then
                            dt_WaitHandle.Rows(i)(19) = "CommissionSheet_Load.aspx?pFormNo=" + dt_WaitHandle.Rows(i)(0).ToString + _
                                                                                 "&pFormSno=" + dt_WaitHandle.Rows(i)(1).ToString + _
                                                                                 "&pWID=" + dt_WaitHandle.Rows(i)(17).ToString + _
                                                                                 "&pBEndTime=" + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyyMMddHHmm") + _
                                                                                 "&pAEndTime=" + CDate(dt_WaitHandle.Rows(i)(16)).ToString("yyyyMMddHHmm")
                        Else
                            dt_WaitHandle.Rows(i)(19) = "CommissionSheet_Load.aspx?pFormNo=" + dt_WaitHandle.Rows(i)(0).ToString + _
                                                                                 "&pFormSno=" + dt_WaitHandle.Rows(i)(1).ToString + _
                                                                                 "&pWID=" + dt_WaitHandle.Rows(i)(17).ToString + _
                                                                                 "&pBEndTime=" + CDate(dt_WaitHandle.Rows(i)(14)).ToString("yyyyMMddHHmm") + _
                                                                                 "&pAEndTime=" + ""
                        End If
                    End If
                Next
            End If
            GridView1.DataSource = dt_WaitHandle
            GridView1.DataBind()
            '
            'Modify-End
        End If
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pIdx As Integer, ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        '-----------------------------------------------------------------
        '-- 開發委託
        '-----------------------------------------------------------------
        '----基本-------------------------------------------------    
        'BUYER
        If pFieldName = "APPBUYER" Then
            DAPPBUYER.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAPPBUYER.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='700' and DKey='BUYER' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAPPBUYER.Items.Add(ListItem1)
                Next
            End If
        End If
        '用途
        If pFieldName = "USAGE" Then
            DUSAGE.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUSAGE.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='USAGE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUSAGE.Items.Add(ListItem1)
                Next
            End If
        End If
        '需圖面
        If pFieldName = "NEEDMAP" Then
            DNEEDMAP.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNEEDMAP.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='NEEDMAP' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNEEDMAP.Items.Add(ListItem1)
                Next
            End If
        End If
        '----樣品-------------------------------------------------    
        '製品區分
        If pFieldName = "PRO" Then
            DPRO.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPRO.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='PRO' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPRO.Items.Add(ListItem1)
                Next
            End If
        End If
        '長度單位(企)
        If pFieldName = "PLENUN" Then
            DPLENUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPLENUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='LENUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPLENUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '數量單位(企)
        If pFieldName = "PQTYUN" Then
            DPQTYUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPQTYUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='QTYUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPQTYUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '長度單位(EA)
        If pFieldName = "EALENUN" Then
            DEALENUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEALENUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='LENUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEALENUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '數量單位(EA)
        If pFieldName = "EAQTYUN" Then
            DEAQTYUN.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEAQTYUN.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='QTYUN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEAQTYUN.Items.Add(ListItem1)
                Next
            End If
        End If
        '----開發-------------------------------------------------    
        '型別
        If pFieldName = "SIZENO" Then
            DSIZENO.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSIZENO.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='SIZE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSIZENO.Items.Add(ListItem1)
                Next
            End If
        End If
        '鏈條形式
        If pFieldName = "ITEM" Then
            DITEM.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DITEM.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='CHAIN' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DITEM.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶
        If pFieldName = "TATYPE" Then
            DTATYPE.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTATYPE.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='TAPE' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTATYPE.Items.Add(ListItem1)
                Next
            End If
        End If
        '鏈齒顏色-SEL
        If pFieldName = "ECOLSEL" Then
            DECOLSEL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DECOLSEL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-A' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DECOLSEL.Items.Add(ListItem1)
                Next
            End If
        End If
        '丸紐-SEL
        If pFieldName = "CCOLSEL" Then
            DCCOLSEL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCCOLSEL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-A' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCCOLSEL.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶-色番(同)
        If pFieldName = "TACOL" Then
            DTACOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTACOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTACOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶-色番(左)
        If pFieldName = "TALCOL" Then
            DTALCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTALCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTALCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '布帶-色番(右)
        If pFieldName = "TARCOL" Then
            DTARCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTARCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTARCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫上-色番(同)
        If pFieldName = "THUPCOL" Then
            DTHUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫上-色番(左)
        If pFieldName = "THLUPCOL" Then
            DTHLUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫上-色番(左右)
        If pFieldName = "THLRUPCOL" Then
            DTHLRUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLRUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLRUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If


        '縫上-色番(右)
        If pFieldName = "THRUPCOL" Then
            DTHRUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫上-色番(右右)
        If pFieldName = "THRRUPCOL" Then
            DTHRRUPCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRRUPCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRRUPCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫下-色番(同)
        If pFieldName = "THLOCOL" Then
            DTHLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '縫下-色番(左)
        If pFieldName = "THLLOCOL" Then
            DTHLLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫下-色番(左右)
        If pFieldName = "THLRLOCOL" Then
            DTHLRLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHLRLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHLRLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫下-色番(右)
        If pFieldName = "THRLOCOL" Then
            DTHRLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        '縫下-色番(右右)
        If pFieldName = "THRRLOCOL" Then
            DTHRRLOCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTHRRLOCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTHRRLOCOL.Items.Add(ListItem1)
                Next
            End If
        End If

        'X-色番
        If pFieldName = "XMCOL" Then
            DXMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DXMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DXMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'A-色番
        If pFieldName = "AMCOL" Then
            DAMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'B-色番
        If pFieldName = "BMCOL" Then
            DBMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'C-色番
        If pFieldName = "CMCOL" Then
            DCMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'D-色番
        If pFieldName = "DMCOL" Then
            DDMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'E-色番
        If pFieldName = "EMCOL" Then
            DEMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'F-色番
        If pFieldName = "FMCOL" Then
            DFMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'G-色番
        If pFieldName = "GMCOL" Then
            DGMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DGMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DGMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        'H-色番
        If pFieldName = "HMCOL" Then
            DHMCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DHMCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DHMCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '緯紗-色番
        If pFieldName = "LYCOL" Then
            DLYCOL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLYCOL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='COLOR-B' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLYCOL.Items.Add(ListItem1)
                Next
            End If
        End If
        '製圖者
        If pFieldName = "MAKEMAP" Then
            DMAKEMAP.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMAKEMAP.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='MAKEMAP' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMAKEMAP.Items.Add(ListItem1)
                Next
            End If
        End If
        '難易度
        If pFieldName = "LEVEL" Then
            DLEVEL.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLEVEL.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='LEVEL' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLEVEL.Items.Add(ListItem1)
                Next
            End If
        End If
        '-----------------------------------------------------------------
        '-- 製造委託
        '-----------------------------------------------------------------
        'OP1
        '----遲納原因類別1
        If pFieldName = "OP1DELAYC1" Then
            DOP1DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP1DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP1DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP1DELAYC2" Then
            DOP1DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP1DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP1DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP2
        '----遲納原因類別1
        If pFieldName = "OP2DELAYC1" Then
            DOP2DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP2DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP2DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP2DELAYC2" Then
            DOP2DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP2DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP2DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP3
        '----遲納原因類別1
        If pFieldName = "OP3DELAYC1" Then
            DOP3DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP3DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP3DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP3DELAYC2" Then
            DOP3DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP3DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP3DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP4
        '----遲納原因類別1
        If pFieldName = "OP4DELAYC1" Then
            DOP4DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP4DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP4DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP4DELAYC2" Then
            DOP4DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP4DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP4DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP5
        '----遲納原因類別1
        If pFieldName = "OP5DELAYC1" Then
            DOP5DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP5DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP5DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP5DELAYC2" Then
            DOP5DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP5DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP5DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP6
        '----遲納原因類別1
        If pFieldName = "OP6DELAYC1" Then
            DOP6DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP6DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP6DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP6DELAYC2" Then
            DOP6DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP6DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP6DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP7
        '----遲納原因類別1
        If pFieldName = "OP7DELAYC1" Then
            DOP7DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP7DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP7DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP7DELAYC2" Then
            DOP7DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP7DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP7DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        'OP8
        '----遲納原因類別1
        If pFieldName = "OP8DELAYC1" Then
            DOP8DELAYC1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP8DELAYC1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP8DELAYC1.Items.Add(ListItem1)
                Next
            End If
        End If
        '----遲納原因類別2
        If pFieldName = "OP8DELAYC2" Then
            DOP8DELAYC2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DOP8DELAYC2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='DELAYC2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DOP8DELAYC2.Items.Add(ListItem1)
                Next
            End If
        End If
        '-----------------------------------------------------------------
        '-- 開發見本
        '-----------------------------------------------------------------
        '承認-作成者
        If pFieldName = "WF1" Then
            D3WF1.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF1.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF1.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-責任者
        If pFieldName = "WF2" Then
            D3WF2.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF2.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF2.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造1
        If pFieldName = "WF3NAME" Then
            D3WF3Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF3Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF3Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF3" Then
            D3WF3.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF3.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF3.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造2
        If pFieldName = "WF4NAME" Then
            D3WF4Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF4Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF4Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF4" Then
            D3WF4.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF4.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF4.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造3
        If pFieldName = "WF5NAME" Then
            D3WF5Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF5Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF5Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF5" Then
            D3WF5.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF5.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF5.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-製造4
        If pFieldName = "WF6NAME" Then
            D3WF6Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF6Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF6Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF6" Then
            D3WF6.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF6.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF6.Items.Add(ListItem1)
                Next
            End If
        End If
        '承認-廠長
        If pFieldName = "WF7NAME" Then
            D3WF7Name.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF7Name.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF7Name.Items.Add(ListItem1)
                Next
            End If
        End If
        If pFieldName = "WF7" Then
            D3WF7.Items.Clear()
            If pIdx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    D3WF7.Items.Add(ListItem1)
                End If
            Else
                sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF7.Items.Add(ListItem1)
                Next
            End If
        End If
        '-----------------------------------------------------------------
        '-- 原單純
        '-----------------------------------------------------------------
        '目前不控制
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     延遲處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim Pos As Integer = InStr(e.Row.Cells(7).Text.ToString, "],")
            '
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'Modify-Start by Joy  2012/7/31 新納期對應
            '
            'Dim Str1 As String = Mid(e.Row.Cells(7).Text.ToString, 1, Pos)
            'Dim Str2 As String = Mid(e.Row.Cells(7).Text.ToString, Pos + 3, Len(e.Row.Cells(7).Text.ToString))
            'e.Row.Cells(7).Text = Str1 + "<br/>" + Str2
            If Pos > 0 Then
                Dim Str1 As String = Mid(e.Row.Cells(7).Text.ToString, 1, Pos)
                Dim Str2 As String = Mid(e.Row.Cells(7).Text.ToString, Pos + 3, Len(e.Row.Cells(7).Text.ToString))
                e.Row.Cells(7).Text = Str1 + "<br/>" + Str2
            End If
            '
            'Modify-End
        End If
    End Sub

    '*****************************************************************
    '**
    '**     設定ImageButton的圖檔
    '**
    '*****************************************************************
    Protected Sub SetImageButtonImageFile(ByVal pIndex As Integer)
        MultiView1.ActiveViewIndex = pIndex
        '設定頁次影像
        If pIndex = 0 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Blank.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 1 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Blank.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 2 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Blank.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 3 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Blank.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Button.jpg"
        End If
        If pIndex = 4 Then
            ImageButton1.ImageUrl = "Images/DevelopmentCommission_Button.jpg"
            ImageButton2.ImageUrl = "Images/ManufactureCommission_Button.jpg"
            ImageButton3.ImageUrl = "Images/DevelopmentSample_Button.jpg"
            ImageButton4.ImageUrl = "Images/DevelopmentGentani_Button.jpg"
            ImageButton5.ImageUrl = "Images/SignHistory_Blank.jpg"
        End If
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton1 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        SetImageButtonImageFile(0)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton2 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton2.Click
        SetImageButtonImageFile(1)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton3 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton3_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton3.Click
        SetImageButtonImageFile(2)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton4 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton4.Click
        SetImageButtonImageFile(3)
    End Sub
    '*****************************************************************
    '**
    '**     ImageButton5 事件處理程序
    '**
    '*****************************************************************
    Protected Sub ImageButton5_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton5.Click
        SetImageButtonImageFile(4)
    End Sub
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'Add-Start by Joy  2012/7/31 新納期對應
    '
    '*****************************************************************
    '**(GetFlowLoading)
    '**     取得負荷狀態
    '**
    '*****************************************************************
    Function GetFlowLoading(ByVal pStep As Integer) As Integer
        Dim RtnCode As Integer = 0
        Dim xKey As String = ""
        Dim sql As String
        '
        Select Case pStep
            Case 40
                xKey = "OPLOAD-" & HOP1.Text
            Case 50
                xKey = "OPLOAD-" & HOP2.Text
            Case 60
                xKey = "OPLOAD-" & HOP3.Text
            Case 70
                xKey = "OPLOAD-" & HOP4.Text
            Case 80
                xKey = "OPLOAD-" & HOP5.Text
            Case 90
                xKey = "OPLOAD-" & HOP6.Text
            Case 100
                xKey = "OPLOAD-" & HOP7.Text
            Case Else
                xKey = "OPLOAD-" & HOP8.Text
        End Select
        '
        sql = "Select * From M_Referp "
        sql &= "Where Cat  = '2002' "
        sql &= "  And DKey = '" & xKey & "' "
        Dim dt_Referp = uDataBase.GetDataTable(sql)
        If dt_Referp.Rows.Count > 0 Then
            RtnCode = 1
        End If
        '
        Return RtnCode
    End Function
    'Add-End

End Class
