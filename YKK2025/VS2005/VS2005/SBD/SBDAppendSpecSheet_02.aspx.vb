Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDAppendSpecSheet_02
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

        Response.Cookies("PGM").Value = "SBDAppendSpecSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
    End Sub





    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDAppendSpecPath")



        SQL = "Select * From F_SBDAppendSpecSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Value = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DBuyer.Value = DBAdapter1.Rows(0).Item("Buyer")
            DSupplier.Value = DBAdapter1.Rows(0).Item("Supplier") '外注廠商
            DVendor.Value = DBAdapter1.Rows(0).Item("Vendor")              'Vendor

            DSurfSheetNo.Value = DBAdapter1.Rows(0).Item("SurfSheetNo") ' 新表面處理NO
            DSurfSupplier.Value = DBAdapter1.Rows(0).Item("SurfSupplier") '外注廠商

            DCap.Value = DBAdapter1.Rows(0).Item("Cap")              '日產能





            DSchedule.Value = DBAdapter1.Rows(0).Item("Schedule")              '基礎日程

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


            '檢測結果
            DQCResult.Text = DBAdapter1.Rows(0).Item("QCResult")

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



            If Trim(DBAdapter1.Rows(0).Item("FinalSampleFile")) <> "" Then
                LFinalSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("FinalSampleFile")  '切結書
                LFinalSampleFile.Visible = True
            Else
                LFinalSampleFile.Visible = False
            End If


        End If








    End Sub


    Protected Sub DNo_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.ServerChange

    End Sub

    Protected Sub DDivision_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.TextChanged

    End Sub
End Class

