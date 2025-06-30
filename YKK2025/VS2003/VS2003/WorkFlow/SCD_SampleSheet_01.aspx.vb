Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SCD_SampleSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSampleSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSampleSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents DDEVPRD As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDEVNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTAWIDTH As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSAMPLEFILE As System.Web.UI.WebControls.Image
    Protected WithEvents DCODENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSIZENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTACOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DECOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTHCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSAMPLEFILE As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DTNLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCNITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTNRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCSITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCDITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DOP1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF5 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DTALINE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAPPBUYER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDATE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF3Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF4Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF5Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF6Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF7Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BRNO As System.Web.UI.WebControls.Button
    Protected WithEvents DRNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO2ITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO1ITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents D1Other As System.Web.UI.WebControls.Label
    Protected WithEvents D2Other As System.Web.UI.WebControls.Label
    Protected WithEvents LQCFILE1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQCFILE1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFILE3 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFILE2 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFILE4 As System.Web.UI.HtmlControls.HtmlInputFile
    
    Protected WithEvents Hidden1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Hidden2 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents DQCFILE5 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LQCFILE5 As System.Web.UI.WebControls.HyperLink

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String = "CL"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

    Dim wUserName As String = ""    '姓名代理用

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示
            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
            End If
            SetPopupFunction()      '設定彈出視窗事件
        Else
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查
            ShowMessage()           '上傳資料檢查及顯示訊息
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("CommissionNo").Value = ""    'CodeNo, DevelopDataPicker使用

        Response.Cookies("PGM").Value = "SCD_SampleSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""

        '實際樣品
        If DSAMPLEFILE.Visible Then
            If DSAMPLEFILE.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "實際樣品"
                Else
                    Message = Message & ", " & "實際樣品"
                End If
            End If
        End If

        'update by alin
        '品測報告1
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品測報告"
                Else
                    Message = Message & ", " & "品測報告"
                End If
            End If
        End If
        '品測報告2
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品測報告"
                Else
                    Message = Message & ", " & "品測報告"
                End If
            End If
        End If
        '品測報告3
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品測報告"
                Else
                    Message = Message & ", " & "品測報告"
                End If
            End If
        End If
        '品測報告4
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品測報告"
                Else
                    Message = Message & ", " & "品測報告"
                End If
            End If
        End If

        '品測報告5
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "品測報告"
                Else
                    Message = Message & ", " & "品測報告"
                End If
            End If
        End If
        '--------------------------------------------------


        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SampleFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_SampleSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SampleSheet")
        If DBDataSet1.Tables("F_SampleSheet").Rows.Count > 0 Then
            '表單資料
            DRNO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Rno")                   'Commission-No
            DNo.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("No")                     'No
            DAPPBUYER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("AppBuyer")         'Customer
            DDATE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Date")                 '發行日
            DSIZENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SizeNo")             'Size
            DITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Item")                 'Item
            DCODENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CodeNo")             'Code No
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile") <> "" Then          '實際樣品
                LSAMPLEFILE.ImageUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile")
            Else
                LSAMPLEFILE.Visible = False
            End If
            DTAWIDTH.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TAWidth")           '布帶寬度
            DDEVNO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevNo")               '開發No
            DDEVPRD.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevPrd")             '開發期間
            DTACOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TACol")               '布帶Color
            DTALINE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TALine")             '條紋線Color
            DECOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("ECol")                 '務齒
            DCCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CCol")                 '丸紐
            DTHCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("THCol")               '縫工線
            DOTHER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other")               '其他

            'update by alin
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1") <> "" Then              '品測報告1
                LQCFILE1.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1")
            Else
                LQCFILE1.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2") <> "" Then              '品測報告2
                LQCFILE2.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2")
            Else
                LQCFILE2.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3") <> "" Then              '品測報告3
                LQCFILE3.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3")
            Else
                LQCFILE3.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4") <> "" Then              '品測報告4
                LQCFILE4.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4")
            Else
                LQCFILE4.Visible = False
            End If

            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5") <> "" Then              '品測報告5
                LQCFILE5.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5")
            Else
                LQCFILE5.Visible = False
            End If

            D1Other.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other1")           'Other1
            Hidden1.Value = D1Other.Text
            D2Other.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other2")           'Other2
            Hidden2.Value = D2Other.Text
            DO1ITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O1Item")           'O1Item
            DO2ITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O2Item")           'O2Item

            '----------------------------------

            DTNLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNLItem")           'TAPE NAT-左
            DTNRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNRItem")           'TAPE NAT-右
            DTSLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSLItem")           'TAPE SET-左
            DTSRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSRItem")           'TAPE SET-右
            DTDLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDLItem")           'TAPE DYED-左
            DTDRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDRItem")           'TAPE DYED-右
            DCNITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CNItem")             'CHAIN NAT
            DCSITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CSItem")             'CHAIN SET
            DCDITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CDItem")             'CHAIN DYED
            DCITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CItem")               'CORD
            DOP1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP1")                   '工程1
            DOP2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP2")                   '工程2
            DOP3.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP3")                   '工程3
            DOP4.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP4")                   '工程4
            DOP5.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP5")                   '工程5
            DOP6.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP6")                   '工程6
            SetFieldData("WF1", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF1"))          '承認WF1
            SetFieldData("WF2", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF2"))          '承認WF2
            SetFieldData("WF3", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3"))          '承認WF3
            SetFieldData("WF4", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4"))          '承認WF4
            SetFieldData("WF5", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5"))          '承認WF5
            SetFieldData("WF6", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6"))          '承認WF6
            SetFieldData("WF7", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7"))          '承認WF7
            SetFieldData("WF3NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3Name"))      '承認者部門WF3
            SetFieldData("WF4NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4Name"))      '承認者部門WF4
            SetFieldData("WF5NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5Name"))      '承認者部門WF5
            SetFieldData("WF6NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6Name"))      '承認者部門WF6
            SetFieldData("WF7NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7Name"))      '承認者部門WF7
            '交易資料
            DBDataSet1.Clear()
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '說明

                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '延遲理由代碼
                        If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '延遲理由
                            DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '延遲其他說明
                        End If

                    End If
                End If
            End If
            '核定履歷資料
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            SQL = SQL + "Order by Unique_ID Desc "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet9, "DecideHistory")
            DataGrid9.DataSource = DBDataSet9
            DataGrid9.DataBind()
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     延遲理由代碼點選後事件
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            '電子簽章未使用
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Attach") = 1 Then
                'DRefMapFile.Visible = True
                'DMapFile.Visible = True
            Else
                'DRefMapFile.Visible = False
                'DMapFile.Visible = False
            End If
            '儲存按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '新
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
                '--
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc1")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                '新
                BNG1.Attributes("onclick") = "this.disabled = true;" & "Button('NG1', '" + BNG1.Value + "');"
                '--
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc2")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                '新
                BNG2.Attributes("onclick") = "this.disabled = true;" & "Button('NG2', '" + BNG2.Value + "');"
                '--
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                '新
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '" + BOK.Value + "');"
                '--
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If

            '遲納管理
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet顯示
                DSampleSheet1.Visible = True     '表單Sheet-1
                DSampleSheet2.Visible = True     '表單Sheet-2
                DDescSheet.Visible = True        '說明Sheet

                '遲納管理
                If wDelay = 1 Then
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                    Else
                        DDelay.Visible = False  '延遲Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '延遲理由代碼
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                    DReason.Visible = True         '延遲理由
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                    DReasonDesc.Visible = True     '延遲其他說明
                    Top = 1285
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 1157
                End If

                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                '連結顯示---需再修改
                LSAMPLEFILE.Visible = True     'Sample File

                'update by alin
                LQCFILE1.Visible = True         'QC File1
                LQCFILE2.Visible = True         'QC File2
                LQCFILE3.Visible = True         'QC File3
                LQCFILE4.Visible = True         'QC File4
                LQCFILE5.Visible = True         'QC File5
                '---------------------------

                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                '核定履歷
                DHistoryLabel.Style.Add("Top", Top + 24)  '核定履歷
                DataGrid9.Style.Add("Top", Top + 48)     '核定履歷
            End If
        Else
            Top = 1085
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '連結隱藏
            LSAMPLEFILE.Visible = False    'Sample File

            'update by alin
            LQCFILE1.Visible = False        'QC File1
            LQCFILE2.Visible = False        'QC File2
            LQCFILE3.Visible = False        'QC File3
            LQCFILE4.Visible = False        'QC File4
            LQCFILE5.Visible = False        'QC File5
            '------------------------------

            '按鈕位置
            BNG1.Style.Add("Top", Top)     'NG1按鈕
            BNG2.Style.Add("Top", Top)     'NG2按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕

        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        '資料選擇
        BRNO.Attributes("onclick") = "DataPicker('Develop');"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        '按鈕及RequestedField的Top位置
        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 1157
                Else
                    If DDelay.Visible = True Then
                        Top = 1285
                    Else
                        Top = 1157
                    End If
                End If
            End If
        Else
            Top = 1085
        End If
        '----
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性

        SetFieldAttribute(pPost)     '表單各欄位屬性及欄位輸入檢查等設定
    End Sub
    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        '-Commission-No特別處理-------------------------------------------------------------------
        If pPost = "New" Then
            DRNO.Text = ""
            DRNO.ForeColor = Color.White
            DRNO.BackColor = Color.White
        End If
        '------------------------------------------------------------------------------------------
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                'DNo.BackColor = Color.LightGray
                DNo.BackColor = Color.White
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '修改+檢查
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "異常：需輸入Ｎｏ")
                DNo.Visible = True
            Case 2  '修改
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""
        'AppBuyer
        Select Case FindFieldInf("AppBuyer")
            Case 0  '顯示
                DAPPBUYER.BackColor = Color.LightGray
                DAPPBUYER.ReadOnly = True
                DAPPBUYER.Visible = True
            Case 1  '修改+檢查
                DAPPBUYER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppBuyerRqd", "DAppBuyer", "異常：需輸入Customer")
                DAPPBUYER.Visible = True
            Case 2  '修改
                DAPPBUYER.BackColor = Color.Yellow
                DAPPBUYER.Visible = True
            Case Else   '隱藏
                DAPPBUYER.Visible = False
        End Select
        If pPost = "New" Then DAPPBUYER.Text = ""
        '發行日
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDATE.BackColor = Color.LightGray
                DDATE.ReadOnly = True
                DDATE.Visible = True
            Case 1  '修改+檢查
                DDATE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入發行日")
                DDATE.Visible = True
            Case 2  '修改
                DDATE.BackColor = Color.Yellow
                DDATE.Visible = True
            Case Else   '隱藏
                DDATE.Visible = False
        End Select
        If pPost = "New" Then DDATE.Text = CStr(DateTime.Now.Today)
        'Size
        Select Case FindFieldInf("SizeNo")
            Case 0  '顯示
                DSIZENO.BackColor = Color.LightGray
                DSIZENO.ReadOnly = True
                DSIZENO.Visible = True
            Case 1  '修改+檢查
                DSIZENO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSizeNoRqd", "DSizeNo", "異常：需輸入Size")
                DSIZENO.Visible = True
            Case 2  '修改
                DSIZENO.BackColor = Color.Yellow
                DSIZENO.Visible = True
            Case Else   '隱藏
                DSIZENO.Visible = False
        End Select
        If pPost = "New" Then DSIZENO.Text = ""
        'Item
        Select Case FindFieldInf("Item")
            Case 0  '顯示
                DITEM.BackColor = Color.LightGray
                DITEM.ReadOnly = True
                DITEM.Visible = True
            Case 1  '修改+檢查
                DITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DItemRqd", "DItem", "異常：需輸入Item")
                DITEM.Visible = True
            Case 2  '修改
                DITEM.BackColor = Color.Yellow
                DITEM.Visible = True
            Case Else   '隱藏
                DITEM.Visible = False
        End Select
        If pPost = "New" Then DITEM.Text = ""
        'Tape
        Select Case FindFieldInf("CodeNo")
            Case 0  '顯示
                DCODENO.BackColor = Color.LightGray
                DCODENO.ReadOnly = True
                DCODENO.Visible = True
            Case 1  '修改+檢查
                DCODENO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeNoRqd", "DCodeNo", "異常：需輸入Tape")
                DCODENO.Visible = True
            Case 2  '修改
                DCODENO.BackColor = Color.Yellow
                DCODENO.Visible = True
            Case Else   '隱藏
                DCODENO.Visible = False
        End Select
        If pPost = "New" Then DCODENO.Text = ""
        '實際樣品
        Select Case FindFieldInf("SampleFile")
            Case 0  '顯示
                DSAMPLEFILE.Visible = False
                DSAMPLEFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "異常：需輸入實際樣品檔")
                DSAMPLEFILE.Visible = True
                DSAMPLEFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSAMPLEFILE.Visible = True
                DSAMPLEFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSAMPLEFILE.Visible = False
        End Select
        '布帶寬度
        Select Case FindFieldInf("TAWidth")
            Case 0  '顯示
                DTAWIDTH.BackColor = Color.LightGray
                DTAWIDTH.ReadOnly = True
                DTAWIDTH.Visible = True
            Case 1  '修改+檢查
                DTAWIDTH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTAWidthRqd", "DTAWidth", "異常：需輸入布帶寬度")
                DTAWIDTH.Visible = True
            Case 2  '修改
                DTAWIDTH.BackColor = Color.Yellow
                DTAWIDTH.Visible = True
            Case Else   '隱藏
                DTAWIDTH.Visible = False
        End Select
        If pPost = "New" Then DTAWIDTH.Text = ""
        '開發No
        Select Case FindFieldInf("DevNo")
            Case 0  '顯示
                DDEVNO.BackColor = Color.LightGray
                DDEVNO.ReadOnly = True
                DDEVNO.Visible = True
            Case 1  '修改+檢查
                DDEVNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevNoRqd", "DDevNo", "異常：需輸入開發No")
                DDEVNO.Visible = True
            Case 2  '修改
                DDEVNO.BackColor = Color.Yellow
                DDEVNO.Visible = True
            Case Else   '隱藏
                DDEVNO.Visible = False
        End Select
        If pPost = "New" Then DDEVNO.Text = ""
        '開發期間
        Select Case FindFieldInf("DevPrd")
            Case 0  '顯示
                DDEVPRD.BackColor = Color.LightGray
                DDEVPRD.ReadOnly = True
                DDEVPRD.Visible = True
            Case 1  '修改+檢查
                DDEVPRD.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevPrdRqd", "DDevPrd", "異常：需輸入開發期間")
                DDEVPRD.Visible = True
            Case 2  '修改
                DDEVPRD.BackColor = Color.Yellow
                DDEVPRD.Visible = True
            Case Else   '隱藏
                DDEVPRD.Visible = False
        End Select
        If pPost = "New" Then DDEVPRD.Text = ""
        '布帶
        Select Case FindFieldInf("TACol")
            Case 0  '顯示
                DTACOL.BackColor = Color.LightGray
                DTACOL.ReadOnly = True
                DTACOL.Visible = True
            Case 1  '修改+檢查
                DTACOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTAColRqd", "DTACol", "異常：需輸入布帶")
                DTACOL.Visible = True
            Case 2  '修改
                DTACOL.BackColor = Color.Yellow
                DTACOL.Visible = True
            Case Else   '隱藏
                DTACOL.Visible = False
        End Select
        If pPost = "New" Then DTACOL.Text = ""
        '條紋線
        Select Case FindFieldInf("TALine")
            Case 0  '顯示
                DTALINE.BackColor = Color.LightGray
                DTALINE.ReadOnly = True
                DTALINE.Visible = True
            Case 1  '修改+檢查
                DTALINE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTALineRqd", "DTALine", "異常：需輸入條紋線")
                DTALINE.Visible = True
            Case 2  '修改
                DTALINE.BackColor = Color.Yellow
                DTALINE.Visible = True
            Case Else   '隱藏
                DTALINE.Visible = False
        End Select
        If pPost = "New" Then DTALINE.Text = ""
        '務齒
        Select Case FindFieldInf("ECol")
            Case 0  '顯示
                DECOL.BackColor = Color.LightGray
                DECOL.ReadOnly = True
                DECOL.Visible = True
            Case 1  '修改+檢查
                DECOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEColRqd", "DECol", "異常：需輸入務齒")
                DECOL.Visible = True
            Case 2  '修改
                DECOL.BackColor = Color.Yellow
                DECOL.Visible = True
            Case Else   '隱藏
                DECOL.Visible = False
        End Select
        If pPost = "New" Then DECOL.Text = ""
        '丸紐
        Select Case FindFieldInf("CCol")
            Case 0  '顯示
                DCCOL.BackColor = Color.LightGray
                DCCOL.ReadOnly = True
                DCCOL.Visible = True
            Case 1  '修改+檢查
                DCCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCColRqd", "DCCol", "異常：需輸入丸紐")
                DCCOL.Visible = True
            Case 2  '修改
                DCCOL.BackColor = Color.Yellow
                DCCOL.Visible = True
            Case Else   '隱藏
                DCCOL.Visible = False
        End Select
        If pPost = "New" Then DCCOL.Text = ""
        '縫工線
        Select Case FindFieldInf("THCol")
            Case 0  '顯示
                DTHCOL.BackColor = Color.LightGray
                DTHCOL.ReadOnly = True
                DTHCOL.Visible = True
            Case 1  '修改+檢查
                DTHCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHColRqd", "DTHCol", "異常：需輸入縫工線")
                DTHCOL.Visible = True
            Case 2  '修改
                DTHCOL.BackColor = Color.Yellow
                DTHCOL.Visible = True
            Case Else   '隱藏
                DTHCOL.Visible = False
        End Select
        If pPost = "New" Then DTHCOL.Text = ""
        '其他
        Select Case FindFieldInf("Other")
            Case 0  '顯示
                DOTHER.BackColor = Color.LightGray
                DOTHER.ReadOnly = True
                DOTHER.Visible = True
            Case 1  '修改+檢查
                DOTHER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOtherRqd", "DOther", "異常：需輸入其他")
                DOTHER.Visible = True
            Case 2  '修改
                DOTHER.BackColor = Color.Yellow
                DOTHER.Visible = True
            Case Else   '隱藏
                DOTHER.Visible = False
        End Select
        If pPost = "New" Then DOTHER.Text = ""

        'update by alin
        '品測報告1
        Select Case FindFieldInf("QCFile1")
            Case 0  '顯示
                DQCFILE1.Visible = False
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFileRqd1", "DQCFile1", "異常：需輸入品測報告檔1")
                DQCFILE1.Visible = True
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCFILE1.Visible = True
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFILE1.Visible = False
        End Select
        '品測報告2
        Select Case FindFieldInf("QCFile2")
            Case 0  '顯示
                DQCFILE2.Visible = False
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFileRqd2", "DQCFile2", "異常：需輸入品測報告檔2")
                DQCFILE2.Visible = True
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCFILE2.Visible = True
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFILE2.Visible = False
        End Select
        '品測報告3
        Select Case FindFieldInf("QCFile3")
            Case 0  '顯示
                DQCFILE3.Visible = False
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFileRqd3", "DQCFile3", "異常：需輸入品測報告檔3")
                DQCFILE3.Visible = True
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCFILE3.Visible = True
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFILE3.Visible = False
        End Select
        '品測報告4
        Select Case FindFieldInf("QCFile4")
            Case 0  '顯示
                DQCFILE4.Visible = False
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFileRqd4", "DQCFile4", "異常：需輸入品測報告檔4")
                DQCFILE4.Visible = True
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCFILE4.Visible = True
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFILE4.Visible = False
        End Select

        '品測報告5
        Select Case FindFieldInf("QCFile5")
            Case 0  '顯示
                DQCFILE5.Visible = False
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCFileRqd5", "DQCFile5", "異常：需輸入品測報告檔4")
                DQCFILE5.Visible = True
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCFILE5.Visible = True
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCFILE5.Visible = False
        End Select
        '--------------------------------

        'Tape Nat-左
        Select Case FindFieldInf("TNLItem")
            Case 0  '顯示
                DTNLITEM.BackColor = Color.LightGray
                DTNLITEM.ReadOnly = True
                DTNLITEM.Visible = True
            Case 1  '修改+檢查
                DTNLITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTNLItemRqd", "DTNLItem", "異常：需輸入Tape Nat-左")
                DTNLITEM.Visible = True
            Case 2  '修改
                DTNLITEM.BackColor = Color.Yellow
                DTNLITEM.Visible = True
            Case Else   '隱藏
                DTNLITEM.Visible = False
        End Select
        If pPost = "New" Then DTNLITEM.Text = ""
        'Tape Nat-右
        Select Case FindFieldInf("TNRItem")
            Case 0  '顯示
                DTNRITEM.BackColor = Color.LightGray
                DTNRITEM.ReadOnly = True
                DTNRITEM.Visible = True
            Case 1  '修改+檢查
                DTNRITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTNRItemRqd", "DTNRItem", "異常：需輸入Tape Nat-右")
                DTNRITEM.Visible = True
            Case 2  '修改
                DTNRITEM.BackColor = Color.Yellow
                DTNRITEM.Visible = True
            Case Else   '隱藏
                DTNRITEM.Visible = False
        End Select
        If pPost = "New" Then DTNRITEM.Text = ""
        'Tape Set-左
        Select Case FindFieldInf("TSLItem")
            Case 0  '顯示
                DTSLITEM.BackColor = Color.LightGray
                DTSLITEM.ReadOnly = True
                DTSLITEM.Visible = True
            Case 1  '修改+檢查
                DTSLITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTSLItemRqd", "DTSLItem", "異常：需輸入Tape Set-左")
                DTSLITEM.Visible = True
            Case 2  '修改
                DTSLITEM.BackColor = Color.Yellow
                DTSLITEM.Visible = True
            Case Else   '隱藏
                DTSLITEM.Visible = False
        End Select
        If pPost = "New" Then DTSLITEM.Text = ""
        'Tape Set-右
        Select Case FindFieldInf("TSRItem")
            Case 0  '顯示
                DTSRITEM.BackColor = Color.LightGray
                DTSRITEM.ReadOnly = True
                DTSRITEM.Visible = True
            Case 1  '修改+檢查
                DTSRITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTSRItemRqd", "DTSRItem", "異常：需輸入Tape Set-右")
                DTSRITEM.Visible = True
            Case 2  '修改
                DTSRITEM.BackColor = Color.Yellow
                DTSRITEM.Visible = True
            Case Else   '隱藏
                DTSRITEM.Visible = False
        End Select
        If pPost = "New" Then DTSRITEM.Text = ""
        'Tape Dyed-左
        Select Case FindFieldInf("TDLItem")
            Case 0  '顯示
                DTDLITEM.BackColor = Color.LightGray
                DTDLITEM.ReadOnly = True
                DTDLITEM.Visible = True
            Case 1  '修改+檢查
                DTDLITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTDLItemRqd", "DTDLItem", "異常：需輸入Tape Dyed-左")
                DTDLITEM.Visible = True
            Case 2  '修改
                DTDLITEM.BackColor = Color.Yellow
                DTDLITEM.Visible = True
            Case Else   '隱藏
                DTDLITEM.Visible = False
        End Select
        If pPost = "New" Then DTDLITEM.Text = ""
        'Tape Dyed-右
        Select Case FindFieldInf("TDRItem")
            Case 0  '顯示
                DTDRITEM.BackColor = Color.LightGray
                DTDRITEM.ReadOnly = True
                DTDRITEM.Visible = True
            Case 1  '修改+檢查
                DTDRITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTDRItemRqd", "DTDRItem", "異常：需輸入Tape Dyed-右")
                DTDRITEM.Visible = True
            Case 2  '修改
                DTDRITEM.BackColor = Color.Yellow
                DTDRITEM.Visible = True
            Case Else   '隱藏
                DTDRITEM.Visible = False
        End Select
        If pPost = "New" Then DTDRITEM.Text = ""
        'Chain Nat
        Select Case FindFieldInf("CNItem")
            Case 0  '顯示
                DCNITEM.BackColor = Color.LightGray
                DCNITEM.ReadOnly = True
                DCNITEM.Visible = True
            Case 1  '修改+檢查
                DCNITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCNItemRqd", "DCNItem", "異常：需輸入Chain Nat")
                DCNITEM.Visible = True
            Case 2  '修改
                DCNITEM.BackColor = Color.Yellow
                DCNITEM.Visible = True
            Case Else   '隱藏
                DCNITEM.Visible = False
        End Select
        If pPost = "New" Then DCNITEM.Text = ""
        'Chain Set
        Select Case FindFieldInf("CSItem")
            Case 0  '顯示
                DCSITEM.BackColor = Color.LightGray
                DCSITEM.ReadOnly = True
                DCSITEM.Visible = True
            Case 1  '修改+檢查
                DCSITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCSItemRqd", "DCSItem", "異常：需輸入Chain Set")
                DCSITEM.Visible = True
            Case 2  '修改
                DCSITEM.BackColor = Color.Yellow
                DCSITEM.Visible = True
            Case Else   '隱藏
                DCSITEM.Visible = False
        End Select
        If pPost = "New" Then DCSITEM.Text = ""
        'Chain Dyed
        Select Case FindFieldInf("CDItem")
            Case 0  '顯示
                DCDITEM.BackColor = Color.LightGray
                DCDITEM.ReadOnly = True
                DCDITEM.Visible = True
            Case 1  '修改+檢查
                DCDITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCDItemRqd", "DCDItem", "異常：需輸入Chain Dyed")
                DCDITEM.Visible = True
            Case 2  '修改
                DCDITEM.BackColor = Color.Yellow
                DCDITEM.Visible = True
            Case Else   '隱藏
                DCDITEM.Visible = False
        End Select

        'update by alin
        'O1ITEM
        Select Case FindFieldInf("O1Item")
            Case 0  '顯示
                DO1ITEM.BackColor = Color.LightGray
                DO1ITEM.ReadOnly = True
                DO1ITEM.Visible = True
            Case 1  '修改+檢查
                DO1ITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DO1ItemRqd", "DO1Item", "異常：需輸入Other Item1")
                DO1ITEM.Visible = True
            Case 2  '修改
                DO1ITEM.BackColor = Color.Yellow
                DO1ITEM.Visible = True
            Case Else   '隱藏
                DO1ITEM.Visible = False
        End Select

        'O2ITEM
        Select Case FindFieldInf("O2Item")
            Case 0  '顯示
                DO2ITEM.BackColor = Color.LightGray
                DO2ITEM.ReadOnly = True
                DO2ITEM.Visible = True
            Case 1  '修改+檢查
                DO2ITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DO2ItemRqd", "DO2Item", "異常：需輸入Other Item2")
                DO2ITEM.Visible = True
            Case 2  '修改
                DO2ITEM.BackColor = Color.Yellow
                DO2ITEM.Visible = True
            Case Else   '隱藏
                DO2ITEM.Visible = False
        End Select
        '----------------------------------


        If pPost = "New" Then DCDITEM.Text = ""
        'Cord
        Select Case FindFieldInf("CItem")
            Case 0  '顯示
                DCITEM.BackColor = Color.LightGray
                DCITEM.ReadOnly = True
                DCITEM.Visible = True
            Case 1  '修改+檢查
                DCITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCItemRqd", "DCItem", "異常：需輸入Cord")
                DCITEM.Visible = True
            Case 2  '修改
                DCITEM.BackColor = Color.Yellow
                DCITEM.Visible = True
            Case Else   '隱藏
                DCITEM.Visible = False
        End Select
        If pPost = "New" Then DCITEM.Text = ""
        '工程１
        Select Case FindFieldInf("OP1")
            Case 0  '顯示
                DOP1.BackColor = Color.LightGray
                DOP1.ReadOnly = True
                DOP1.Visible = True
            Case 1  '修改+檢查
                DOP1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP1Rqd", "DOP1", "異常：需輸入工程１")
                DOP1.Visible = True
            Case 2  '修改
                DOP1.BackColor = Color.Yellow
                DOP1.Visible = True
            Case Else   '隱藏
                DOP1.Visible = False
        End Select
        If pPost = "New" Then DOP1.Text = ""
        '工程２
        Select Case FindFieldInf("OP2")
            Case 0  '顯示
                DOP2.BackColor = Color.LightGray
                DOP2.ReadOnly = True
                DOP2.Visible = True
            Case 1  '修改+檢查
                DOP2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP2Rqd", "DOP2", "異常：需輸入工程２")
                DOP2.Visible = True
            Case 2  '修改
                DOP2.BackColor = Color.Yellow
                DOP2.Visible = True
            Case Else   '隱藏
                DOP2.Visible = False
        End Select
        If pPost = "New" Then DOP2.Text = ""
        '工程３
        Select Case FindFieldInf("OP3")
            Case 0  '顯示
                DOP3.BackColor = Color.LightGray
                DOP3.ReadOnly = True
                DOP3.Visible = True
            Case 1  '修改+檢查
                DOP3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP3Rqd", "DOP3", "異常：需輸入工程３")
                DOP3.Visible = True
            Case 2  '修改
                DOP3.BackColor = Color.Yellow
                DOP3.Visible = True
            Case Else   '隱藏
                DOP3.Visible = False
        End Select
        If pPost = "New" Then DOP3.Text = ""
        '工程４
        Select Case FindFieldInf("OP4")
            Case 0  '顯示
                DOP4.BackColor = Color.LightGray
                DOP4.ReadOnly = True
                DOP4.Visible = True
            Case 1  '修改+檢查
                DOP4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP4Rqd", "DOP4", "異常：需輸入工程４")
                DOP4.Visible = True
            Case 2  '修改
                DOP4.BackColor = Color.Yellow
                DOP4.Visible = True
            Case Else   '隱藏
                DOP4.Visible = False
        End Select
        If pPost = "New" Then DOP4.Text = ""
        '工程５
        Select Case FindFieldInf("OP5")
            Case 0  '顯示
                DOP5.BackColor = Color.LightGray
                DOP5.ReadOnly = True
                DOP5.Visible = True
            Case 1  '修改+檢查
                DOP5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP5Rqd", "DOP5", "異常：需輸入工程５")
                DOP5.Visible = True
            Case 2  '修改
                DOP5.BackColor = Color.Yellow
                DOP5.Visible = True
            Case Else   '隱藏
                DOP5.Visible = False
        End Select
        If pPost = "New" Then DOP5.Text = ""
        '工程６
        Select Case FindFieldInf("OP6")
            Case 0  '顯示
                DOP6.BackColor = Color.LightGray
                DOP6.ReadOnly = True
                DOP6.Visible = True
            Case 1  '修改+檢查
                DOP6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP6Rqd", "DOP6", "異常：需輸入工程６")
                DOP6.Visible = True
            Case 2  '修改
                DOP6.BackColor = Color.Yellow
                DOP6.Visible = True
            Case Else   '隱藏
                DOP6.Visible = False
        End Select
        If pPost = "New" Then DOP6.Text = ""
        '作成者
        Select Case FindFieldInf("WF1")
            Case 0  '顯示
                DWF1.BackColor = Color.LightGray
                DWF1.Visible = True
            Case 1  '修改+檢查
                DWF1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF1Rqd", "DWF1", "異常：需輸入作成者")
                DWF1.Visible = True
            Case 2  '修改
                DWF1.BackColor = Color.Yellow
                DWF1.Visible = True
            Case Else   '隱藏
                DWF1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF1", "ZZZZZZ")
        'EA責任者
        Select Case FindFieldInf("WF2")
            Case 0  '顯示
                DWF2.BackColor = Color.LightGray
                DWF2.Visible = True
            Case 1  '修改+檢查
                DWF2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF2Rqd", "DWF2", "異常：需輸入EA責任者")
                DWF2.Visible = True
            Case 2  '修改
                DWF2.BackColor = Color.Yellow
                DWF2.Visible = True
            Case Else   '隱藏
                DWF2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF2", "ZZZZZZ")
        '製造１
        Select Case FindFieldInf("WF3")
            Case 0  '顯示
                DWF3.BackColor = Color.LightGray
                DWF3.Visible = True
            Case 1  '修改+檢查
                DWF3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF3Rqd", "DWF3", "異常：需輸入製造１")
                DWF3.Visible = True
            Case 2  '修改
                DWF3.BackColor = Color.Yellow
                DWF3.Visible = True
            Case Else   '隱藏
                DWF3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF3", "ZZZZZZ")
        '製造２
        Select Case FindFieldInf("WF4")
            Case 0  '顯示
                DWF4.BackColor = Color.LightGray
                DWF4.Visible = True
            Case 1  '修改+檢查
                DWF4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF4Rqd", "DWF4", "異常：需輸入製造２")
                DWF4.Visible = True
            Case 2  '修改
                DWF4.BackColor = Color.Yellow
                DWF4.Visible = True
            Case Else   '隱藏
                DWF4.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF4", "ZZZZZZ")
        '製造３
        Select Case FindFieldInf("WF5")
            Case 0  '顯示
                DWF5.BackColor = Color.LightGray
                DWF5.Visible = True
            Case 1  '修改+檢查
                DWF5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF5Rqd", "DWF5", "異常：需輸入製造３")
                DWF5.Visible = True
            Case 2  '修改
                DWF5.BackColor = Color.Yellow
                DWF5.Visible = True
            Case Else   '隱藏
                DWF5.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF5", "ZZZZZZ")
        '製造４
        Select Case FindFieldInf("WF6")
            Case 0  '顯示
                DWF6.BackColor = Color.LightGray
                DWF6.Visible = True
            Case 1  '修改+檢查
                DWF6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF6Rqd", "DWF6", "異常：需輸入製造４")
                DWF6.Visible = True
            Case 2  '修改
                DWF6.BackColor = Color.Yellow
                DWF6.Visible = True
            Case Else   '隱藏
                DWF6.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF6", "ZZZZZZ")
        '廠長
        Select Case FindFieldInf("WF7")
            Case 0  '顯示
                DWF7.BackColor = Color.LightGray
                DWF7.Visible = True
            Case 1  '修改+檢查
                DWF7.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF7Rqd", "DWF7", "異常：需輸入廠長")
                DWF7.Visible = True
            Case 2  '修改
                DWF7.BackColor = Color.Yellow
                DWF7.Visible = True
            Case Else   '隱藏
                DWF7.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF7", "ZZZZZZ")
        '製造１承認者部門
        Select Case FindFieldInf("WF3Name")
            Case 0  '顯示
                DWF3Name.BackColor = Color.LightGray
                DWF3Name.Visible = True
            Case 1  '修改+檢查
                DWF3Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF3NameRqd", "DWF3Name", "異常：需輸入製造１部門")
                DWF3Name.Visible = True
            Case 2  '修改
                DWF3Name.BackColor = Color.Yellow
                DWF3Name.Visible = True
            Case Else   '隱藏
                DWF3Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF3NAME", "ZZZZZZ")
        '製造２承認者部門
        Select Case FindFieldInf("WF4Name")
            Case 0  '顯示
                DWF4Name.BackColor = Color.LightGray
                DWF4Name.Visible = True
            Case 1  '修改+檢查
                DWF4Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF4NameRqd", "DWF4Name", "異常：需輸入製造２部門")
                DWF4Name.Visible = True
            Case 2  '修改
                DWF4Name.BackColor = Color.Yellow
                DWF4Name.Visible = True
            Case Else   '隱藏
                DWF4Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF4NAME", "ZZZZZZ")
        '製造３承認者部門
        Select Case FindFieldInf("WF5Name")
            Case 0  '顯示
                DWF5Name.BackColor = Color.LightGray
                DWF5Name.Visible = True
            Case 1  '修改+檢查
                DWF5Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF5NameRqd", "DWF5Name", "異常：需輸入製造３部門")
                DWF5Name.Visible = True
            Case 2  '修改
                DWF5Name.BackColor = Color.Yellow
                DWF5Name.Visible = True
            Case Else   '隱藏
                DWF5Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF5NAME", "ZZZZZZ")
        '製造４承認者部門
        Select Case FindFieldInf("WF6Name")
            Case 0  '顯示
                DWF6Name.BackColor = Color.LightGray
                DWF6Name.Visible = True
            Case 1  '修改+檢查
                DWF6Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF6NameRqd", "DWF6Name", "異常：需輸入製造４部門")
                DWF6Name.Visible = True
            Case 2  '修改
                DWF6Name.BackColor = Color.Yellow
                DWF6Name.Visible = True
            Case Else   '隱藏
                DWF6Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF6NAME", "ZZZZZZ")
        '廠長承認者部門
        Select Case FindFieldInf("WF7Name")
            Case 0  '顯示
                DWF7Name.BackColor = Color.LightGray
                DWF7Name.Visible = True
            Case 1  '修改+檢查
                DWF7Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF7NameRqd", "DWF7Name", "異常：需輸入廠長部門")
                DWF7Name.Visible = True
            Case 2  '修改
                DWF7Name.BackColor = Color.Yellow
                DWF7Name.Visible = True
            Case Else   '隱藏
                DWF7Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF7NAME", "ZZZZZZ")

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
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        idx = FindFieldInf(pFieldName)

        '作成者
        If pFieldName = "WF1" Then
            DWF1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = LTrim(RTrim(DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")))
                    ListItem1.Value = LTrim(RTrim(DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")))
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF1.Items.Add(ListItem1)
                End If
            End If
        End If
        'EA責任者
        If pFieldName = "WF2" Then
            DWF2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF2.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造１
        If pFieldName = "WF3" Then
            DWF3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF3.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF3.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造２
        If pFieldName = "WF4" Then
            DWF4.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF4.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF4.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造３
        If pFieldName = "WF5" Then
            DWF5.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF5.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF5.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造４
        If pFieldName = "WF6" Then
            DWF6.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF6.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF6.Items.Add(ListItem1)
                Next
            End If
        End If
        '廠長
        If pFieldName = "WF7" Then
            DWF7.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF7.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF7.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造１部門
        If pFieldName = "WF3NAME" Then
            DWF3Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF3Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF3Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造２部門
        If pFieldName = "WF4NAME" Then
            DWF4Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF4Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF4Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造３部門
        If pFieldName = "WF5NAME" Then
            DWF5Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF5Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF5Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '製造４部門
        If pFieldName = "WF6NAME" Then
            DWF6Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF6Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF6Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '廠長部門
        If pFieldName = "WF7NAME" Then
            DWF7Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF7Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF7Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由代碼
        If pFieldName = "ReasonCode" Then
            DReasonCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReasonCode.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DKey")
                    ListItem1.Value = DBTable1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由
        If pFieldName = "Reason" Then
            SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                DReason.Text = DBTable1.Rows(i).Item("Data")
            Next
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FindFieldInf)
    '**     尋找表單欄位屬性
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i <= 60 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
                Dim Message As String = ""

                'Check理由
                If ErrCode = 0 Then
                    If DReasonCode.Visible = True Then
                        If DReasonCode.SelectedValue = "99" Then
                            If DReasonDesc.Text = "" Then ErrCode = 9210
                        End If
                    End If
                End If
                '儲存資料
                If ErrCode = 0 Then
                    ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
                    ModifyTranData("SAVE", "0")       '更新交易資料
                Else
                    If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '上傳檔案ErrCode=0

                If ErrCode = 0 Then
                    Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                        "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                    Response.Redirect(URL)
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("OK", 0, "1")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("NG1", 1, "2")
            End If
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")
            End If
        End If
    End Sub
    'Joy-090113
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '取得此關卡各按鈕的Data Status

        'Check工程核定
        If ErrCode = 0 Then
            If DWF1.SelectedValue = "" Then ErrCode = 9010
            If DWF2.SelectedValue = "" Then ErrCode = 9010

            If DWF3Name.SelectedValue <> "" Then
                If DWF3.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF4Name.SelectedValue <> "" Then
                If DWF4.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF5Name.SelectedValue <> "" Then
                If DWF5.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF6Name.SelectedValue <> "" Then
                If DWF6.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF7Name.SelectedValue <> "" Then
                If DWF7.SelectedValue = "" Then ErrCode = 9010
            End If

            If DWF3.SelectedValue <> "" Then
                If DWF3Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF4.SelectedValue <> "" Then
                If DWF4Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF5.SelectedValue <> "" Then
                If DWF5Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF6.SelectedValue <> "" Then
                If DWF6Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF7.SelectedValue <> "" Then
                If DWF7Name.SelectedValue = "" Then ErrCode = 9010
            End If
        End If
        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If
        '檢查委託書No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("002001", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '難易度

            While Run = True
                Run = False     '執行Flag=不執行
                '--取得下一關參數設定---------
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=通知
                Dim pCount As Integer
                '--取得LeadTime參數設定---------
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--取得工程負荷參數設定---------
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--其他參數設定---------
                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                        '取得表單流水號
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End

                            '設定委託No
                            DNo.Text = SetNo(NewFormSno)
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            ModifyTranData(pFun, pSts)

                            '流程資料結束
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo, Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                        End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then
                    '找出簽核者User ID
                    Dim wAllocateID As String = ""
                    Dim SQL As String = ""
                    Dim DBDataSet1 As New DataSet
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                    OleDbConnection1.Open()

                    If pAction = 0 Then
                        SQL = "Select UserID From M_Users "
                        SQL = SQL & " Where Active = 1 "
                        Select Case wStep
                            Case 1, 500
                                SQL = SQL & "   And UserName =  '" & DWF2.SelectedValue & "'"
                            Case 10
                                SQL = SQL & "   And UserName =  '" & DWF3.SelectedValue & "'"
                            Case 20
                                SQL = SQL & "   And UserName =  '" & DWF4.SelectedValue & "'"
                            Case 30
                                SQL = SQL & "   And UserName =  '" & DWF5.SelectedValue & "'"
                            Case 40
                                SQL = SQL & "   And UserName =  '" & DWF6.SelectedValue & "'"
                            Case 50
                                SQL = SQL & "   And UserName =  '" & DWF7.SelectedValue & "'"
                            Case Else
                        End Select
                        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                        DBAdapter1.Fill(DBDataSet1, "M_Users")
                        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")
                        OleDbConnection1.Close()
                        '如果為無效使用者時 Action=2 (直接跳999結束)
                        If wAllocateID = "" Then pAction = 2
                    End If
                    '取得簽核者
                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                    If ErrCode = 0 Then pAction = 0
                End If

                '建置流程資料
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        wNextGate = ""
                        For i = 1 To pCount
                            '取得下一關人員(訊息時使用)
                            If wNextGate = "" Then
                                wNextGate = pNextGate(i)
                            Else
                                wNextGate = "," & pNextGate(i)
                            End If

                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '取得核定者-群組行是曆
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '取得工程負荷最後日時
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數

                            '取得預定開始,完成日程計算
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆

                            '建置流程資料
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        'Modify-End
                    End If
                End If
                '儲存表單資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If
                    Else
                        If pNextStep = 999 Then     '工程結束嗎?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                        Else
                            ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                    End If
                    '工程結束(999)並OK按鈕時處理
                    If pNextStep = 999 Then
                        If pFun = "OK" Then
                            InterfaceProc()
                        End If
                    End If

                    If pFlowType <> 3 Then
                        MultiJob = 0
                    Else
                        MultiJob = 1
                    End If

                    If (pRunNextStep = 1) And (pFlowType = 0 Or pFlowType = 3) Then
                        Run = True
                        RepeatRun = True
                        pAction = 0
                    End If

                    wStep = pNextStep     '下一工程關卡號碼
                    wFormSno = NewFormSno '下一工程表單流水號
                Else
                    EnabledButton()   '起動Button運作
                    If ErrCode = 9110 Then Message = "取得表單流水號計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9120 Then Message = "流程資料更新異常,請確認或連絡系統人員!"
                    If ErrCode = 9130 Then Message = "下工程計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9131 Then Message = "無下工程管理人,請確認或連絡系統人員!"
                    If ErrCode = 9140 Then Message = "工程預定開始及完成日計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9150 Then Message = "下一工程資料建置異常,請確認或連絡系統人員!"
                    If ErrCode = 9160 Then Message = "工程結束資料建置異常(999),請確認或連絡系統人員!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '儲存表單ErrCode=0
            End While       '重覆執行

            If ErrCode = 0 Then
                '--郵件傳送--------
                oCommon.SendMail()

                Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "承認者指定有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     取得表單進度狀態
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            End If
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SampleFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_SampleSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "            '1~5
        SQl = SQl + "Date, AppBuyer, SizeNo, Item, CodeNo, "               '6~10
        SQl = SQl + "SampleFile, TAWidth, DevNo, DevPrd, TACol, "          '11~15
        SQl = SQl + "TALine, ECol, CCol, THCol, Other, "                   '16~20

        'update by alin
        SQl = SQl + "QCFile1,QCFile2,QCFile3,QCFile4,QCFile5,TNLItem, TNRItem, TSLItem, TSRItem, "         '21~25
        SQl = SQl + "TDLItem, TDRItem, CNITem, CSItem, CDItem, "           '26~30
        SQl = SQl + "CItem, "                                              '31
        SQl = SQl + "OP1, OP2, OP3, OP4, OP5, OP6, "                       '32~37
        SQl = SQl + "WF1, WF2, WF3, WF4, WF5 , WF6, WF7, "                 '38~44
        SQl = SQl + "WF3Name, WF4Name, WF5Name, WF6Name, WF7Name, "        '45~49
        SQl = SQl + "Rno, "                                                '50
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime ,Other1,Other2,O1Item,O2Item"      '51~54
        '------------------------

        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                        '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "                      '結案日
        SQl = SQl + " '002001', "                                   '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "                 '表單流水號
        SQl = SQl + " '" + DNo.Text + "', "                         'NO
        '6~10
        SQl = SQl + " '" + DDATE.Text + "', "                       '發行日期
        SQl = SQl + " '" + DAPPBUYER.Text + "', "                   '客戶
        SQl = SQl + " '" + DSIZENO.Text + "', "                     'Size
        SQl = SQl + " '" + DITEM.Text + "', "                       'Item
        SQl = SQl + " '" + DCODENO.Text + "', "                     'Code-No
        '11~15
        FileName = ""
        If DSAMPLEFILE.Visible Then
            If DSAMPLEFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSAMPLEFILE.PostedFile.FileName, InStr(StrReverse(DSAMPLEFILE.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DSAMPLEFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'Sample-File

        SQl = SQl + " '" + DTAWIDTH.Text + "', "                    '布帶寬度
        SQl = SQl + " '" + DDEVNO.Text + "', "                      '開發No
        SQl = SQl + " '" + DDEVPRD.Text + "', "                     '開發期間
        SQl = SQl + " '" + DTACOL.Text + "', "                      '布帶
        '16~20
        SQl = SQl + " '" + DTALINE.Text + "', "                     '條紋線
        SQl = SQl + " '" + DECOL.Text + "', "                       '務齒
        SQl = SQl + " '" + DCCOL.Text + "', "                       '丸紐
        SQl = SQl + " '" + DTHCOL.Text + "', "                      '縫工線
        SQl = SQl + " N'" + DOTHER.Text + "', "                     '其他
        '21~25

        'update by alin
        '品測檔案1
        FileName = ""
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & Right(DQCFILE1.PostedFile.FileName, InStr(StrReverse(DQCFILE1.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File1

        '品測檔案2
        FileName = ""
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & Right(DQCFILE2.PostedFile.FileName, InStr(StrReverse(DQCFILE2.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE2.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File2

        '品測檔案3
        FileName = ""
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & Right(DQCFILE3.PostedFile.FileName, InStr(StrReverse(DQCFILE3.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE3.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFILE3.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File3

        '品測檔案4
        FileName = ""
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & Right(DQCFILE4.PostedFile.FileName, InStr(StrReverse(DQCFILE4.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE4.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFILE4.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File4

        '品測檔案5
        FileName = ""
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & Right(DQCFILE5.PostedFile.FileName, InStr(StrReverse(DQCFILE5.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE5.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DQCFILE5.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File5
        '------------------------

        SQl = SQl + " '" + DTNLITEM.Text + "', "                    '
        SQl = SQl + " '" + DTNRITEM.Text + "', "                    '
        SQl = SQl + " '" + DTSLITEM.Text + "', "                    '
        SQl = SQl + " '" + DTSRITEM.Text + "', "                    '
        '26~30
        SQl = SQl + " '" + DTDLITEM.Text + "', "                    '
        SQl = SQl + " '" + DTDRITEM.Text + "', "                    '
        SQl = SQl + " '" + DCNITEM.Text + "', "                     '
        SQl = SQl + " '" + DCSITEM.Text + "', "                     '
        SQl = SQl + " '" + DCDITEM.Text + "', "                     '
        '31
        SQl = SQl + " '" + DCITEM.Text + "', "                      '
        '32~37
        SQl = SQl + " '" + DOP1.Text + "', "                        '
        SQl = SQl + " '" + DOP2.Text + "', "                        '
        SQl = SQl + " '" + DOP3.Text + "', "                        '
        SQl = SQl + " '" + DOP4.Text + "', "                        '
        SQl = SQl + " '" + DOP5.Text + "', "                        '
        SQl = SQl + " '" + DOP6.Text + "', "                        '
        '38~44
        SQl = SQl + " '" + DWF1.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF2.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF3.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF4.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF5.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF6.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF7.SelectedValue + "', "               '
        '45~49
        SQl = SQl + " '" + DWF3Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF4Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF5Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF6Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF7Name.SelectedValue + "', "           '
        '50
        SQl = SQl + " '" + DRNO.Text + "', "                    '
        '--------------------------------------------
        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "  '作成者
        SQl = SQl + " '" + NowDateTime + "', "                      '作成時間
        SQl = SQl + " '" + "" + "', "                               '修改者
        SQl = SQl + " '" + NowDateTime + "', "                       '修改時間

        'update by alin
        SQl = SQl + " '" + Hidden1.Value + "', "                       'Other1
        SQl = SQl + " '" + Hidden2.Value + "', "                       'Other2
        SQl = SQl + " '" + DO1ITEM.Text + "', "                       'Item1
        SQl = SQl + " '" + DO2ITEM.Text + "' "                        'Item2
        '------------------------



        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()


        OleDBCommand1.ExecuteNonQuery()




        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SampleFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_SampleSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If

        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDATE.Text & "',"
        SQl = SQl + " AppBuyer = '" & DAPPBUYER.Text & "',"
        SQl = SQl + " SizeNo = '" & DSIZENO.Text & "',"
        SQl = SQl + " Item = '" & DITEM.Text & "',"
        SQl = SQl + " CodeNo = '" & DCODENO.Text & "',"

        If DSAMPLEFILE.Visible Then
            If DSAMPLEFILE.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSAMPLEFILE.PostedFile.FileName, InStr(StrReverse(DSAMPLEFILE.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳草圖
                    DSAMPLEFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " SampleFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " TAWidth = '" & DTAWIDTH.Text & "',"
        SQl = SQl + " DevNo = '" & DDEVNO.Text & "',"
        SQl = SQl + " DevPrd = '" & DDEVPRD.Text & "',"
        SQl = SQl + " TACol = '" & DTACOL.Text & "',"
        SQl = SQl + " TALine = '" & DTALINE.Text & "',"

        SQl = SQl + " ECol = '" & DECOL.Text & "',"
        SQl = SQl + " CCol = '" & DCCOL.Text & "',"
        SQl = SQl + " THCol = '" & DTHCOL.Text & "',"
        SQl = SQl + " Other = N'" & DOTHER.Text & "',"

        'update by alin
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & Right(DQCFILE1.PostedFile.FileName, InStr(StrReverse(DQCFILE1.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE1.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳草圖
                    DQCFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile1 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & Right(DQCFILE2.PostedFile.FileName, InStr(StrReverse(DQCFILE2.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE2.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳草圖
                    DQCFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile2 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & Right(DQCFILE3.PostedFile.FileName, InStr(StrReverse(DQCFILE3.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE3.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳草圖
                    DQCFILE3.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile3 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & Right(DQCFILE4.PostedFile.FileName, InStr(StrReverse(DQCFILE4.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE4.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳草圖
                    DQCFILE4.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile4 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & Right(DQCFILE5.PostedFile.FileName, InStr(StrReverse(DQCFILE5.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE5.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳草圖
                    DQCFILE5.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile5 = N'" & FileName & "',"
            End If
        End If
        '------------------------

        SQl = SQl + " TNLItem = '" & DTNLITEM.Text & "',"
        SQl = SQl + " TNRItem = '" & DTNRITEM.Text & "',"
        SQl = SQl + " TSLItem = '" & DTSLITEM.Text & "',"
        SQl = SQl + " TSRItem = '" & DTSRITEM.Text & "',"
        SQl = SQl + " TDLItem = '" & DTDLITEM.Text & "',"
        SQl = SQl + " TDRItem = '" & DTDRITEM.Text & "',"
        SQl = SQl + " CNITem = '" & DCNITEM.Text & "',"
        SQl = SQl + " CSItem = '" & DCSITEM.Text & "',"
        SQl = SQl + " CDItem = '" & DCDITEM.Text & "',"
        SQl = SQl + " CItem = '" & DCITEM.Text & "',"

        SQl = SQl + " OP1 = '" & DOP1.Text & "',"
        SQl = SQl + " OP2 = '" & DOP2.Text & "',"
        SQl = SQl + " OP3 = '" & DOP3.Text & "',"
        SQl = SQl + " OP4 = '" & DOP4.Text & "',"
        SQl = SQl + " OP5 = '" & DOP5.Text & "',"
        SQl = SQl + " OP6 = '" & DOP6.Text & "',"

        SQl = SQl + " WF1 = '" & DWF1.SelectedValue & "',"
        SQl = SQl + " WF2 = '" & DWF2.SelectedValue & "',"
        SQl = SQl + " WF3 = '" & DWF3.SelectedValue & "',"
        SQl = SQl + " WF4 = '" & DWF4.SelectedValue & "',"
        SQl = SQl + " WF5 = '" & DWF5.SelectedValue & "',"
        SQl = SQl + " WF6 = '" & DWF6.SelectedValue & "',"
        SQl = SQl + " WF7 = '" & DWF7.SelectedValue & "',"

        SQl = SQl + " WF3Name = '" & DWF3Name.SelectedValue & "',"
        SQl = SQl + " WF4Name = '" & DWF4Name.SelectedValue & "',"
        SQl = SQl + " WF5Name = '" & DWF5Name.SelectedValue & "',"
        SQl = SQl + " WF6Name = '" & DWF6Name.SelectedValue & "',"
        SQl = SQl + " WF7Name = '" & DWF7Name.SelectedValue & "',"

        SQl = SQl + " Rno = '" & DRNO.Text & "',"

        'update by alin
        SQl = SQl + " Other1 = '" & Hidden1.Value & "',"
        SQl = SQl + " Other2 = '" & Hidden2.Value & "',"
        SQl = SQl + " O1Item = '" & DO1ITEM.Text & "',"
        SQl = SQl + " O2Item = '" & DO2ITEM.Text & "',"
        '-----------------------------

        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        Dim DBDataSet1 As New DataSet

        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.Open()

        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "T_CommissionNo")

        If DBDataSet1.Tables("T_CommissionNo").Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDBCommand1.ExecuteNonQuery()
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBDataSet1.Tables("T_CommissionNo").Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQl
                    OleDBCommand1.ExecuteNonQuery()
                End If
            End If
        End If

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(InterfaceProc)
    '**     更新系統間I/F資料
    '**
    '*****************************************************************
    Sub InterfaceProc()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SCDSqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        Dim SQl As String
        SQl = "Update Commission Set "
        SQl = SQl + " SAMFN = '" & "新版" & "',"
        SQl = SQl + " SAMFormNo = '" & wFormNo & "',"
        SQl = SQl + " SAMFormSno = '" & CStr(wFormSno) & "',"
        SQl = SQl + " ModOp = '" & "WF" & "',"
        SQl = SQl + " ModDate = '" & NowDateTime & "' "
        SQl = SQl + " Where Rno  =  '" & DRNO.Text & "'"
        SQl = SQl + "   And DevNo =  '" & DDEVNO.Text & "'"
        SQl = SQl + "   And CodeNo =  '" & DCODENO.Text & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編製委託No 
    '**
    '*****************************************************************
    Function SetNo(ByVal Seq As Integer) As String
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim i As Integer

        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 10 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If

        'If UPFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
        'Check上傳檔案格式
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function InputCheck() As Integer
        InputCheck = 0
        'No
        If InputCheck = 0 Then
            If FindFieldInf("No") = 1 Then
                If DNo.Text = "" Then InputCheck = 1
            End If
        End If
        '發行日
        If InputCheck = 0 Then
            If FindFieldInf("Date") = 1 Then
                If DDATE.Text = "" Then InputCheck = 1
            End If
        End If
        'AppBuyer
        If InputCheck = 0 Then
            If FindFieldInf("AppBuyer") = 1 Then
                If DAPPBUYER.Text = "" Then InputCheck = 1
            End If
        End If
        'Size
        If InputCheck = 0 Then
            If FindFieldInf("SizeNo") = 1 Then
                If DSIZENO.Text = "" Then InputCheck = 1
            End If
        End If
        'Item
        If InputCheck = 0 Then
            If FindFieldInf("Item") = 1 Then
                If DITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape
        If InputCheck = 0 Then
            If FindFieldInf("CodeNo") = 1 Then
                If DCODENO.Text = "" Then InputCheck = 1
            End If
        End If
        '實際樣品
        If InputCheck = 0 Then
            If FindFieldInf("SampleFile") = 1 Then
                If DSAMPLEFILE.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '布帶寬度
        If InputCheck = 0 Then
            If FindFieldInf("TAWidth") = 1 Then
                If DTAWIDTH.Text = "" Then InputCheck = 1
            End If
        End If
        '開發No
        If InputCheck = 0 Then
            If FindFieldInf("DevNo") = 1 Then
                If DDEVNO.Text = "" Then InputCheck = 1
            End If
        End If
        '開發期間
        If InputCheck = 0 Then
            If FindFieldInf("DevPrd") = 1 Then
                If DDEVPRD.Text = "" Then InputCheck = 1
            End If
        End If
        '布帶
        If InputCheck = 0 Then
            If FindFieldInf("TACol") = 1 Then
                If DTACOL.Text = "" Then InputCheck = 1
            End If
        End If
        '條紋線
        If InputCheck = 0 Then
            If FindFieldInf("TALine") = 1 Then
                If DTALINE.Text = "" Then InputCheck = 1
            End If
        End If
        '務齒
        If InputCheck = 0 Then
            If FindFieldInf("ECol") = 1 Then
                If DECOL.Text = "" Then InputCheck = 1
            End If
        End If
        '丸紐
        If InputCheck = 0 Then
            If FindFieldInf("CCol") = 1 Then
                If DCCOL.Text = "" Then InputCheck = 1
            End If
        End If
        '縫工線
        If InputCheck = 0 Then
            If FindFieldInf("THCol") = 1 Then
                If DTHCOL.Text = "" Then InputCheck = 1
            End If
        End If
        '其他
        If InputCheck = 0 Then
            If FindFieldInf("Other") = 1 Then
                If DOTHER.Text = "" Then InputCheck = 1
            End If
        End If
        '品測報告1
        If InputCheck = 0 Then
            If FindFieldInf("QCFile1") = 1 Then
                If DQCFILE1.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '品測報告2
        If InputCheck = 0 Then
            If FindFieldInf("QCFile2") = 1 Then
                If DQCFILE2.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '品測報告3
        If InputCheck = 0 Then
            If FindFieldInf("QCFile3") = 1 Then
                If DQCFILE3.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '品測報告4
        If InputCheck = 0 Then
            If FindFieldInf("QCFile4") = 1 Then
                If DQCFILE4.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '品測報告5
        If InputCheck = 0 Then
            If FindFieldInf("QCFile5") = 1 Then
                If DQCFILE5.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        'Tape Nat-左
        If InputCheck = 0 Then
            If FindFieldInf("TNLItem") = 1 Then
                If DTNLITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Nat-右
        If InputCheck = 0 Then
            If FindFieldInf("TNRItem") = 1 Then
                If DTNRITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Set-左
        If InputCheck = 0 Then
            If FindFieldInf("TSLItem") = 1 Then
                If DTSLITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Set-右
        If InputCheck = 0 Then
            If FindFieldInf("TSRItem") = 1 Then
                If DTSRITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Dyed-左
        If InputCheck = 0 Then
            If FindFieldInf("TDLItem") = 1 Then
                If DTDLITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Dyed-右
        If InputCheck = 0 Then
            If FindFieldInf("TDRItem") = 1 Then
                If DTDRITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Chain Nat
        If InputCheck = 0 Then
            If FindFieldInf("CNItem") = 1 Then
                If DCNITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Chain Set
        If InputCheck = 0 Then
            If FindFieldInf("CSItem") = 1 Then
                If DCSITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Chain Dyed
        If InputCheck = 0 Then
            If FindFieldInf("CDItem") = 1 Then
                If DCDITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'O1ITEM
        If InputCheck = 0 Then
            If FindFieldInf("O1Item") = 1 Then
                If DO1ITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'O2ITEM
        If InputCheck = 0 Then
            If FindFieldInf("O2Item") = 1 Then
                If DO2ITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Cord
        If InputCheck = 0 Then
            If FindFieldInf("CItem") = 1 Then
                If DCITEM.Text = "" Then InputCheck = 1
            End If
        End If
        '工程１
        If InputCheck = 0 Then
            If FindFieldInf("OP1") = 1 Then
                If DOP1.Text = "" Then InputCheck = 1
            End If
        End If
        '工程２
        If InputCheck = 0 Then
            If FindFieldInf("OP2") = 1 Then
                If DOP2.Text = "" Then InputCheck = 1
            End If
        End If
        '工程３
        If InputCheck = 0 Then
            If FindFieldInf("OP3") = 1 Then
                If DOP3.Text = "" Then InputCheck = 1
            End If
        End If
        '工程４
        If InputCheck = 0 Then
            If FindFieldInf("OP4") = 1 Then
                If DOP4.Text = "" Then InputCheck = 1
            End If
        End If
        '工程５
        If InputCheck = 0 Then
            If FindFieldInf("OP5") = 1 Then
                If DOP5.Text = "" Then InputCheck = 1
            End If
        End If
        '工程６
        If InputCheck = 0 Then
            If FindFieldInf("OP6") = 1 Then
                If DOP6.Text = "" Then InputCheck = 1
            End If
        End If
        '作成者
        If InputCheck = 0 Then
            If FindFieldInf("WF1") = 1 Then
                If DWF1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'EA責任者
        If InputCheck = 0 Then
            If FindFieldInf("WF2") = 1 Then
                If DWF2.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造１
        If InputCheck = 0 Then
            If FindFieldInf("WF3") = 1 Then
                If DWF3.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造２
        If InputCheck = 0 Then
            If FindFieldInf("WF4") = 1 Then
                If DWF4.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造３
        If InputCheck = 0 Then
            If FindFieldInf("WF5") = 1 Then
                If DWF5.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造４
        If InputCheck = 0 Then
            If FindFieldInf("WF6") = 1 Then
                If DWF6.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '廠長
        If InputCheck = 0 Then
            If FindFieldInf("WF7") = 1 Then
                If DWF7.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造１承認者部門
        If InputCheck = 0 Then
            If FindFieldInf("WF3Name") = 1 Then
                If DWF3Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造２承認者部門
        If InputCheck = 0 Then
            If FindFieldInf("WF4Name") = 1 Then
                If DWF4Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造３承認者部門
        If InputCheck = 0 Then
            If FindFieldInf("WF5Name") = 1 Then
                If DWF5Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '製造４承認者部門
        If InputCheck = 0 Then
            If FindFieldInf("WF6Name") = 1 Then
                If DWF6Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '廠長承認者部門
        If InputCheck = 0 Then
            If FindFieldInf("WF7Name") = 1 Then
                If DWF7Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
