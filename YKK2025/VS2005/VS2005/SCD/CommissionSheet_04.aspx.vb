Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class CommissionSheet_04
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    Dim wUserIP As String = ""
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim result As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"

        SetParameter()                          '設定參數
        '
        If Not IsPostBack Then                  'PostBack
            SetLinkFile()                       '設定連結檔
            ShowFormData()                      '顯示表單資料
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
        LGenFILE1.Visible = False
        '-----------------------------------------------------------------
        '-- 製造委託
        '-----------------------------------------------------------------
       
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
        Response.Cookies("PGM").Value = "CommissionSheet_02.aspx"                                   '程式名
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")                      '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")                            '工程代碼
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
        wFormNo = Request.QueryString("pFormNo")                                                    '表單號碼
        wFormSno = Request.QueryString("pFormSno")                                                  '表單流水號
    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("COMMISSIONFilePath")

        Dim PathOld1 As String = uCommon.GetAppSetting("HttpOldSCD")
        Dim PathOld2 As String = uCommon.GetAppSetting("COMMISSIONFilePathSCD")

        Dim PathOld3 As String = uCommon.GetAppSetting("HttpOldSCD1")
        Dim PathOld4 As String = uCommon.GetAppSetting("NEW-SAMPLE2001FilePath")


        Dim RtnCode As Integer = 0



        'OLD-SCD
        Dim Path_OLDSAMPLE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-SAMPLEFilePath")
        Dim Path_OLDSAMPLEIMAGE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-SAMPLEIMAGEFilePath")
        Dim Path_OLDMANUFACTURE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-MANUFACTUREFilePath")
        Dim Path_OLDOTHER As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-OTHERFilePath")
        Dim Path_OLDSAMPLE2001 As String = uCommon.GetAppSetting("HttpOld1") & uCommon.GetAppSetting("OLD-SAMPLE2001FilePath")
        '
     

        Dim SQL As String
        SQL = "Select * From F_CommissionSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then




            LManufactureSheet.Text = "製造委託書"
            LManufactureSheet.NavigateUrl = "ManufactureSheet_04.aspx?pFormNo=002002&pFormSno=" + CStr(wFormSno)

            LSampleSheet.Text = "開發見本"
            LSampleSheet.NavigateUrl = "SampleSheet_04.aspx?pFormNo=002002&pFormSno=" + CStr(wFormSno)


            LGentaniSheet.Text = "原單位"
            LGentaniSheet.NavigateUrl = "GentaniSheet_04.aspx?pFormNo=002002&pFormSno=" + CStr(wFormSno)

            LBefOPList.Text = "核定履歷"
            LBefOPList.NavigateUrl = "http://10.245.1.10/WorkFlow/BefOPList.aspx?pFormNo=002002&&pFormSno=" + CStr(wFormSno)

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
            SetFieldData(0, "APPBUYER", dtCommissionSheet.Rows(0).Item("APPBUYER"))     'BUYER
            DSellVendor.Text = dtCommissionSheet.Rows(0).Item("SELLVENDOR")             '委託廠商
            DESYQTY.Text = dtCommissionSheet.Rows(0).Item("ESYQTY")                     '預估量
            DEXPDEL.Text = dtCommissionSheet.Rows(0).Item("EXPDEL")                     '希望交期
            DCUSTITEM.Text = dtCommissionSheet.Rows(0).Item("CUSTITEM")                 '客戶ITEM
            SetFieldData(0, "USAGE", dtCommissionSheet.Rows(0).Item("USAGE"))           '用途
            DORNO.Text = dtCommissionSheet.Rows(0).Item("ORNO")                         'OR-NO
            SetFieldData(0, "NEEDMAP", dtCommissionSheet.Rows(0).Item("NEEDMAP"))       '需圖面
            '草圖
            If dtCommissionSheet.Rows(0).Item("MAPREFFILE") <> "" Then
                If wFormSno < 5001 Then
                    LMAPREFFILE.NavigateUrl = Path_OLDOTHER & RTrim(dtCommissionSheet.Rows(0).Item("MAPREFFILE"))
                    LMAPREFFILE.Visible = True
                Else
                    '封存檔還原  20213/16  jessica
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("MAPREFFILE"))

                    LMAPREFFILE.NavigateUrl = Path & RTrim(dtCommissionSheet.Rows(0).Item("MAPREFFILE"))
                    LMAPREFFILE.Visible = True
                End If
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

            SetFieldData(0, "THLUPCOL", dtCommissionSheet.Rows(0).Item("THLUPCOL")) '縫上-色番(左左)
            DTHLUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPCOLNO")             '縫上-色番(左左)
            DTHLUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLUPYCOLNO")           '縫上-YKK(左左)

            SetFieldData(0, "THLRUPCOL", dtCommissionSheet.Rows(0).Item("THLRUPCOL")) '縫上-色番(左左)
            DTHLRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPCOLNO")             '縫上-色番(左右)
            DTHLRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPYCOLNO")           '縫上-YKK(左右)

            SetFieldData(0, "THLRUPCOL", dtCommissionSheet.Rows(0).Item("THLRUPCOL")) '縫上-色番(左右)
            DTHLRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPCOLNO")             '縫上-色番(左右)
            DTHLRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLRUPYCOLNO")           '縫上-YKK(左右)

            SetFieldData(0, "THRUPCOL", dtCommissionSheet.Rows(0).Item("THRUPCOL")) '縫上-色番(右左)
            DTHRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPCOLNO")             '縫上-色番(右左)
            DTHRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRUPYCOLNO")           '縫上-YKK(右左)
            SetFieldData(0, "THRRUPCOL", dtCommissionSheet.Rows(0).Item("THRRUPCOL")) '縫上-色番(右右)
            DTHRRUPCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRUPCOLNO")             '縫上-色番(右右)
            DTHRRUPYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THRRUPYCOLNO")           '縫上-YKK(右右)

            SetFieldData(0, "THLOCOL", dtCommissionSheet.Rows(0).Item("THLOCOL"))   '縫下-色番(同)
            DTHLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOCOLNO")               '縫下-色番(同)
            DTHLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLOYCOLNO")             '縫下-YKK(同)

            SetFieldData(0, "THLLOCOL", dtCommissionSheet.Rows(0).Item("THLLOCOL")) '縫下-色番(左左)
            DTHLLOCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOCOLNO")             '縫下-色番(左左)
            DTHLLOYCOLNO.Text = dtCommissionSheet.Rows(0).Item("THLLOYCOLNO")           '縫下-YKK(左左)

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
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("MAPFILE"))

                LMAPFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MAPFILE")
                LMAPFILE.Visible = True
            End If
            '----其他附件-------------------------------------------------    
            '適用型別檔
            If dtCommissionSheet.Rows(0).Item("FORTYPEFILE") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("FORTYPEFILE"))

                LFORTYPEFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORTYPEFILE")
                LFORTYPEFILE.Visible = True
            End If
            '品質檢測項目檔
            If dtCommissionSheet.Rows(0).Item("QCCHECKFILE") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCCHECKFILE"))

                LQCCHECKFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCCHECKFILE")
                LQCCHECKFILE.Visible = True
            End If

            'QC檢測文件
            If dtCommissionSheet.Rows(0).Item("QCFILE1") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCFILE1"))

                LQCFILE1.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE1")
                LQCFILE1.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE2") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCFILE2"))

                LQCFILE2.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE2")
                LQCFILE2.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE3") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCFILE3"))

                LQCFILE3.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE3")
                LQCFILE3.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE4") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCFILE4"))

                LQCFILE4.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE4")
                LQCFILE4.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE5") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCFILE5"))

                LQCFILE5.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE5")
                LQCFILE5.Visible = True
            End If
            If dtCommissionSheet.Rows(0).Item("QCFILE6") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("QCFILE6"))

                LQCFILE6.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("QCFILE6")
                LQCFILE6.Visible = True
            End If

            If dtCommissionSheet.Rows(0).Item("GenFILE1") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("GenFILE1"))

                LGenFILE1.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("GenFILE1")
                LGenFILE1.Visible = True
            End If

            '客戶切結書
            If dtCommissionSheet.Rows(0).Item("CONTACTFILE") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("CONTACTFILE"))

                LCONTACTFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("CONTACTFILE")
                LCONTACTFILE.Visible = True
            End If
            '樣品確認書
            If dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE"))

                LSAMPLECONFIRMFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("SAMPLECONFIRMFILE")
                LSAMPLECONFIRMFILE.Visible = True
            End If
            '製造授權書
            If dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE"))

                LMANUFAUTHORITYFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("MANUFAUTHORITYFILE")
                LMANUFAUTHORITYFILE.Visible = True
            End If


            DMANUOUTPRICE.Text = dtCommissionSheet.Rows(0).Item("MANUOUTPRICE")                       '外注加工費
            '報價單
            If dtCommissionSheet.Rows(0).Item("FORCASTFILE") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtCommissionSheet.Rows(0).Item("FORCASTFILE"))


                LFORCASTFILE.NavigateUrl = Path & dtCommissionSheet.Rows(0).Item("FORCASTFILE")
                LFORCASTFILE.Visible = True
            End If
        End If
 
       


    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top - 50 & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
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
    End Sub
 

End Class
