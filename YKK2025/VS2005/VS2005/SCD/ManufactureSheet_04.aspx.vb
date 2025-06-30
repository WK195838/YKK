
Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient

Partial Class ManufactureSheet_04
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
    '-----------------------------------------------------------------
    '-- 製造委託
    '-----------------------------------------------------------------
    Sub SetLinkFile()
        LHINTFILE.Visible = False
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

        '
        If wFormSno >= 5001 Then
            LOldMSheet.Visible = False
          
        End If

        'OLD-SCD
        Dim Path_OLDSAMPLE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-SAMPLEFilePath")
        Dim Path_OLDSAMPLEIMAGE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-SAMPLEIMAGEFilePath")
        Dim Path_OLDMANUFACTURE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-MANUFACTUREFilePath")
        Dim Path_OLDOTHER As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-OTHERFilePath")
        Dim Path_OLDSAMPLE2001 As String = uCommon.GetAppSetting("HttpOld1") & uCommon.GetAppSetting("OLD-SAMPLE2001FilePath")

        Dim SQL As String
        SQL = "Select * From F_CommissionSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)

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
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtManufactureSheet.Rows(0).Item("HINTFILE"))

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
            If wFormSno < 5001 Then                '舊式委託                                 
                LOldMSheet.Visible = True
                LOldMSheet.NavigateUrl = Path_OLDMANUFACTURE & dtManufactureSheet.Rows(0).Item("OP8REM")
                DOP8REM.Text = ""
            Else
                DOP8REM.Text = dtManufactureSheet.Rows(0).Item("OP8REM")            'OP8-遲納原因
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
    End Sub


End Class
