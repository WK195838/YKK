Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDSurfaceSheet_02
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
    Dim HolidayList As New List(Of Integer) '用以記錄假日的欄位索引值
    Dim indexList As New List(Of Integer)   '用以記錄不屬於選取月份的欄位索引值
    Dim DateList As New List(Of String)     '用以記錄所選取的一周日期

    ''' <summary>
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        If Not Me.IsPostBack Then   '不是PostBack
            ShowFormData()      '顯示表單資料
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
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

        Response.Cookies("PGM").Value = "SBDSurfaceSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String


        '按鈕及RequestedField的Top位置
        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter1.Rows.Count > 0 Then
                If DBAdapter1.Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 1100
                Else
                  
                End If
            End If
        Else
            Top = 696
        End If
        '----


    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDSufaceFilePath")



        SQL = "Select * From F_SBDSufaceSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer") 'Buyer

            DSellVendor.Text = DBAdapter1.Rows(0).Item("SellVendor")              'Vendor
            DReqDelDate.Value = DBAdapter1.Rows(0).Item("ReqDelDate")              '希望交期
            DReqQty.Text = DBAdapter1.Rows(0).Item("ReqQty")              '希望交期
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              '品名
            DAttachSample.Text = DBAdapter1.Rows(0).Item("AttachSample")

            DORNO.Text = DBAdapter1.Rows(0).Item("ORNO")              'ORNO

            If Mid(DBAdapter1.Rows(0).Item("OrderTime").ToString, 1, 4) = "1900" Then
                DOrderTime.Value = ""
            Else
                DOrderTime.Value = DBAdapter1.Rows(0).Item("OrderTime")               '下單時間
            End If


            DDevReason.Text = DBAdapter1.Rows(0).Item("DevReason")              '開發理由
            DColor.Text = DBAdapter1.Rows(0).Item("Color")              'color code
            DSampleQty.Text = DBAdapter1.Rows(0).Item("SampleQty") '附樣
         
            DEnglishName.Text = DBAdapter1.Rows(0).Item("EnglishName")              '英文名稱
            DSupplier.Text = DBAdapter1.Rows(0).Item("Supplier") '外注廠商
            DAllowSample.Text = DBAdapter1.Rows(0).Item("AllowSample") '限度樣品

            DCap.Text = DBAdapter1.Rows(0).Item("Cap")              '日產能

            If Mid(DBAdapter1.Rows(0).Item("BFinalDate").ToString, 1, 4) = "1900" Then
                DBFinalDate.Value = ""
            Else
                DBFinalDate.Value = DBAdapter1.Rows(0).Item("BFinalDate")                '預定完成日
            End If




            DSchedule.Text = DBAdapter1.Rows(0).Item("Schedule")              '基礎日程

            If Trim(DBAdapter1.Rows(0).Item("QCReqFile")) <> "" Then
                LQCReqFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QCReqFile")  '品質依賴書
                LQCReqFile.Visible = True
            Else
                LQCReqFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("QAFile")) <> "" Then
                LQAFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QAFile")  '品測報告書
                LQAFile.Visible = True
            Else
                LQAFile.Visible = False
            End If


            If Mid(DBAdapter1.Rows(0).Item("QCDate").ToString, 1, 4) = "1900" Then
                DQCDate.Value = ""
            Else
                DQCDate.Value = DBAdapter1.Rows(0).Item("QCDate")          '品質判定日期
            End If



            DQCResult.Text = DBAdapter1.Rows(0).Item("QCResult") '檢測結果



            DQCRemark.Text = DBAdapter1.Rows(0).Item("QCRemark")              '品質備註


            If Trim(DBAdapter1.Rows(0).Item("ManufFlowFile")) <> "" Then
                LManufFlowFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ManufFlowFile")  '製造流程
                LManufFlowFile.Visible = True
            Else
                LManufFlowFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("OPManualFile")) <> "" Then
                LOPManualFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("OPManualFile")  '作業流程
                LOPManualFile.Visible = True
            Else
                LOPManualFile.Visible = False
            End If


            If Trim(DBAdapter1.Rows(0).Item("ForcastFile")) <> "" Then
                LForcastFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ForcastFile")  ' 報價單
                LForcastFile.Visible = True
            Else
                LForcastFile.Visible = False
            End If


            If Trim(DBAdapter1.Rows(0).Item("ContactFile")) <> "" Then
                LContactFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContactFile")  '切結書
                LContactFile.Visible = True
            Else
                LContactFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("CustSampleFile")) <> "" Then
                LCustSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("CustSampleFile")  '切結書
                LCustSampleFile.Visible = True
            Else
                LCustSampleFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("FinalSampleFile")) <> "" Then
                LFinalSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("FinalSampleFile")  '切結書
                LFinalSampleFile.Visible = True
            Else
                LFinalSampleFile.Visible = False
            End If




           
        End If



      

        'DB連結關閉
        'OleDbConnection1.Close()

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
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


End Class

