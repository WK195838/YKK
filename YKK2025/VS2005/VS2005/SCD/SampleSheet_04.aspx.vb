Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class SampleSheet_04
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

        '-----------------------------------------------------------------
        '-- 開發見本
        '-----------------------------------------------------------------
        Sql = "Select * From FS_SampleSheet "
        Sql &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
        Dim dtSampleSheet As DataTable = uDataBase.GetDataTable(Sql)
        If dtSampleSheet.Rows.Count > 0 Then
            '----基本欄位設定-------------------------------------------------
            D3APPBUYER.Text = dtSampleSheet.Rows(0).Item("AppBuyer")                 'Customer
            If wFormSno < 5001 And dtSampleSheet.Rows(0).Item("AppBuyer") = "" Then  '發行日 
                D3DATE.Text = ""
            Else
                D3DATE.Text = dtSampleSheet.Rows(0).Item("Date")
            End If
            D3SIZENO.Text = dtSampleSheet.Rows(0).Item("SizeNo")                     'Size
            D3ITEM.Text = dtSampleSheet.Rows(0).Item("Item")                         'Item
            D3CODENO.Text = dtSampleSheet.Rows(0).Item("CodeNo")                     'Code No
            '實際樣品
            If dtSampleSheet.Rows(0).Item("SampleFile") <> "" Then
                If wFormSno < 5001 And dtSampleSheet.Rows(0).Item("AppBuyer") = "" Then     '舊式委託
                    LOldSImage.Visible = True
                    LOldSImage.NavigateUrl = Path_OLDSAMPLEIMAGE & dtSampleSheet.Rows(0).Item("SampleFile")
                Else
                    If wFormSno < 5001 Then     '舊式委託(2001)

                        '封存檔還原  20230301  jessica
                        RtnCode = oCommon.RecoveryArchiveFile(PathOld3, PathOld4, dtSampleSheet.Rows(0).Item("SampleFile"))

                        LSAMPLEFILE.ImageUrl = Path_OLDSAMPLE2001 & dtSampleSheet.Rows(0).Item("SampleFile")
                    Else
                        '封存檔還原  20213/16  jessica
                        RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtSampleSheet.Rows(0).Item("SampleFile"))

                        LSAMPLEFILE.ImageUrl = Path & dtSampleSheet.Rows(0).Item("SampleFile")
                    End If
                    LSAMPLEFILE.Visible = True
                End If
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
            If wFormSno < 5001 And dtSampleSheet.Rows(0).Item("AppBuyer") = "" Then     '舊式委託
                LOldSSheet.Visible = True
                LOldSSheet.NavigateUrl = Path_OLDSAMPLE & dtSampleSheet.Rows(0).Item("Other")
                D3OTHER.Text = ""                                                    '其他
            Else
                D3OTHER.Text = dtSampleSheet.Rows(0).Item("Other")                   '其他
            End If
            '----QC附檔-------------------------------------------------
            If dtSampleSheet.Rows(0).Item("QCFile1") <> "" Then                     '品測報告1
                If wFormSno < 5001 Then     '舊式委託(2001)
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld3, PathOld4, dtSampleSheet.Rows(0).Item("QCFile1"))
                    L3QCFILE1.NavigateUrl = Path_OLDSAMPLE2001 & dtSampleSheet.Rows(0).Item("QCFile1")
                Else
                    '封存檔還原  20213/16  jessica
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtSampleSheet.Rows(0).Item("QCFile1"))

                    L3QCFILE1.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile1")
                End If
                L3QCFILE1.Visible = True
            End If
            If dtSampleSheet.Rows(0).Item("QCFile2") <> "" Then                     '品測報告2
                If wFormSno < 5001 Then     '舊式委託(2001)
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld3, PathOld4, dtSampleSheet.Rows(0).Item("QCFile2"))
                    L3QCFILE2.NavigateUrl = Path_OLDSAMPLE2001 & dtSampleSheet.Rows(0).Item("QCFile2")
                Else
                    '封存檔還原  20213/16  jessica
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtSampleSheet.Rows(0).Item("QCFile2"))

                    L3QCFILE2.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile2")
                End If
                L3QCFILE2.Visible = True
            End If
            If dtSampleSheet.Rows(0).Item("QCFile3") <> "" Then                     '品測報告3
                If wFormSno < 5001 Then     '舊式委託(2001)
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld3, PathOld4, dtSampleSheet.Rows(0).Item("QCFile3"))
                    L3QCFILE3.NavigateUrl = Path_OLDSAMPLE2001 & dtSampleSheet.Rows(0).Item("QCFile3")
                Else
                    '封存檔還原  20213/16  jessica
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtSampleSheet.Rows(0).Item("QCFile3"))

                    L3QCFILE3.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile3")
                End If
                L3QCFILE3.Visible = True
            End If
            If dtSampleSheet.Rows(0).Item("QCFile4") <> "" Then                     '品測報告4
                If wFormSno < 5001 Then     '舊式委託(2001)
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld3, PathOld4, dtSampleSheet.Rows(0).Item("QCFile4"))
                    L3QCFILE4.NavigateUrl = Path_OLDSAMPLE2001 & dtSampleSheet.Rows(0).Item("QCFile4")
                Else
                    '封存檔還原  20213/16  jessica
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtSampleSheet.Rows(0).Item("QCFile4"))

                    L3QCFILE4.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile4")
                End If
                L3QCFILE4.Visible = True
            End If
            If dtSampleSheet.Rows(0).Item("QCFile5") <> "" Then                     '品測報告5
                If wFormSno < 5001 Then     '舊式委託(2001)
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld3, PathOld4, dtSampleSheet.Rows(0).Item("QCFile5"))
                    L3QCFILE5.NavigateUrl = Path_OLDSAMPLE2001 & dtSampleSheet.Rows(0).Item("QCFile5")
                Else
                    '封存檔還原  20213/16  jessica
                    RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtSampleSheet.Rows(0).Item("QCFile5"))

                    L3QCFILE5.NavigateUrl = Path & dtSampleSheet.Rows(0).Item("QCFile5")
                End If
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

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetLinkFile)
    '**     設定連結檔
    '**
    '*****************************************************************
    Sub SetLinkFile()
 
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW1' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW2' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOWNAME' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
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
                Sql = "Select * From M_Referp Where Cat='2002' and DKey='FLOW' Order by Data "
                Dim dtFieldData As DataTable = uDataBase.GetDataTable(Sql)
                For i As Integer = 0 To dtFieldData.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtFieldData.Rows(i).Item("Data")
                    ListItem1.Value = dtFieldData.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    D3WF7.Items.Add(ListItem1)
                Next
            End If
        End If

    End Sub

End Class
