Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class MPMProcessesSheet_01
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
    Dim wUserID As String          '使用者ID
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
    Dim AQty, DevNo, Manufout As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件

        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,wUserID) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

            '上傳資料檢查及顯示訊息

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
        If wStep = 40 Or wStep = 500 Then
            Dim SQL As String

            SQL = "Select * From T_waitHandle "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & " And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " and active=1  and step ='" + CStr(wStep) + "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                wUserID = DBAdapter1.Rows(0).Item("Workid")
            End If
        Else
            wUserID = Request.QueryString("pUserID")
        End If



        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "MPMProcessesSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,wUserID)
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, wUserID)
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("MPMProcessesFilePath")



        SQL = "Select * From F_MPMProcessesSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            If wStep = 1 Then
                LNO.Visible = False
            Else
                LNO.Visible = True
                LNO.NavigateUrl = "MPMProcessesReport_02.aspx?&pNo=" + DBAdapter1.Rows(0).Item("No")
            End If


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Value = DBAdapter1.Rows(0).Item("AppDate")              'AppDate           
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DClinter.Text = DBAdapter1.Rows(0).Item("Clinter")              'Clinter
            ' DDivisionCode.SelectedValue = DBAdapter1.Rows(0).Item("DivisionCode") + "-" + DBAdapter1.Rows(0).Item("Division")
            'DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            '  DDivisionCode.Text = DBAdapter1.Rows(0).Item("DivisionCode")              'Division
            SetFieldData("DivisionCode", DBAdapter1.Rows(0).Item("DivisionCode") + "-" + DBAdapter1.Rows(0).Item("Division"))    '類別1

            SetFieldData("Type1", DBAdapter1.Rows(0).Item("Type1"))    '類別1
            SetFieldData("Type2", DBAdapter1.Rows(0).Item("Type2"))    '類別1
            DMapNo.Text = DBAdapter1.Rows(0).Item("MapNo")              'Code
            DQty.Text = DBAdapter1.Rows(0).Item("Qty")              'Qty
            DAQty.Text = DBAdapter1.Rows(0).Item("AQty")              'Qty
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              'MapNo
            DWeight.Text = DBAdapter1.Rows(0).Item("Weight")              'Weight
            DMaterial.Text = DBAdapter1.Rows(0).Item("Material")              'Weight
            DDevNo.Text = DBAdapter1.Rows(0).Item("DevNo")              'Weight
            DManufout.Text = DBAdapter1.Rows(0).Item("ManufOut")              'Weight

            If Mid(DBAdapter1.Rows(0).Item("FinishDate").ToString, 1, 4) = "1900" Then
                DFinishDate.Value = ""
            Else
                DFinishDate.Value = DBAdapter1.Rows(0).Item("FinishDate")               '下單時間
            End If

            '  If Mid(DBAdapter1.Rows(0).Item("AFinishDate").ToString, 1, 4) = "1900" Then
            If wStep = 40 Then
                DAFinishDate.Value = CDate(NowDateTime).ToString("yyyy/MM/dd")
            End If




            If Trim(DBAdapter1.Rows(0).Item("MapFile")) <> "" Then
                LMapFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  '切結書
                LMapFile1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  '品質依賴書
                LMapFile.Visible = True
                LMapFile1.Visible = True
            Else
                LMapFile.Visible = False
                LMapFile1.Visible = False
            End If
        End If


        Dim EngineStr As String

        SQL = "Select Engine,RecDate,StartDate,EndDate,WorkTime,Starter,remark,SeqNo From F_MPMProcessesSheetDT "
        SQL = SQL & " Where delmark =0 and seqno1 =0 and FormSno =  '" & CStr(wFormSno) & "' order by SeqNo"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

        For Each dtr As Data.DataRow In dt.Rows
            EngineStr = "DEngine" + Trim(Str(dtr("SeqNo")))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = dtr("Engine")
        Next



        SQL = "Select * From T_WaitHandle "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
        SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"

        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter2.Rows.Count > 0 Then
            DDecideDesc.Text = DBAdapter2.Rows(0).Item("DecideDesc")       '說明


            If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
                If DDelay.Visible = True Then
                    SetFieldData("ReasonCode", DBAdapter2.Rows(0).Item("ReasonCode"))    '延遲理由代碼
                    If DBAdapter2.Rows(0).Item("ReasonCode") = "" Then
                        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                        DReasonDesc.Text = ""   '延遲其他說明
                    Else
                        DReason.Text = DBAdapter2.Rows(0).Item("Reason")  '延遲理由
                        DReasonDesc.Text = DBAdapter2.Rows(0).Item("ReasonDesc")     '延遲其他說明
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
        Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet9, "DecideHistory")
        GridView2.DataSource = dtWaitHandle1
        GridView2.DataBind()

        'DB連結關閉
        'OleDbConnection1.Close()

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

        'DB連結設定

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)


        If DBAdapter1.Rows.Count > 0 Then
            '電子簽章未使用
            If DBAdapter1.Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBAdapter1.Rows(0).Item("Attach") = 1 Then
            Else
            End If
            '儲存按鈕
            If DBAdapter1.Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBAdapter1.Rows(0).Item("SaveDesc")
                BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBAdapter1.Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBAdapter1.Rows(0).Item("NGDesc1")
                BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                wNGSts1 = DBAdapter1.Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBAdapter1.Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBAdapter1.Rows(0).Item("NGDesc2")
                BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                wNGSts2 = DBAdapter1.Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBAdapter1.Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBAdapter1.Rows(0).Item("OKDesc")
                BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                wOKSts = DBAdapter1.Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If DBAdapter1.Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then

                '遲納管理
                If wDelay = 1 Then
                    If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
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
                    Top = 1100
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 800
                End If



                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If DBAdapter2.Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If

                '按鈕位置
                BNG1.Style.Add("Top", Top)      'NG按鈕
                BNG2.Style.Add("Top", Top)     'NG按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕

            End If
        Else
            Top = 600
            'Sheet隱藏

            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False     '說明
            DDescSheet.Visible = False
            DReasonCode.Visible = False     '延遲理由代碼
            DReason.Visible = False         '延遲理由
            DReasonDesc.Visible = False     '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)      'NG按鈕
            BNG2.Style.Add("Top", Top)     'NG按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕

        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        '日期選擇
        BDate1.Attributes.Add("onClick", "calendarPicker('DFinishDate')")
        BDate2.Attributes.Add("onClick", "calendarPicker('DAFinishDate')")
        BDate3.Attributes.Add("onClick", "calendarPicker('DAppDate')")

        BClienter.Attributes.Add("onClick", "EmpDatePicker('" + wApplyID + "')") '找員工
        'BDate1.Attributes("onclick") = "calendarPicker('Form1.DMapDate');"
        'BDate2.Attributes("onclick") = "calendarPicker('Form1.DSampleDate');"
        'BDate3.Attributes("onclick") = "calendarPicker('Form1.DHalfFinishDdate');"

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
                    If DDelay.Visible = True Then
                        Top = 1600
                    Else
                        Top = 1100
                    End If
                End If
            End If
        Else
            Top = 696
        End If
        '----


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
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
        Dim SQL As String
        Dim Empid As String

        Empid = D1.Text




        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
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





        '日期


        Select Case FindFieldInf("AppDate")
            Case 0  '顯示
                DAppDate.Visible = True
                DAppDate.Style.Add("background-color", "lightgrey")
                DAppDate.Attributes.Add("readonly", "true")
                BDate3.Visible = False

            Case 1  '修改+檢查
                DAppDate.Visible = True
                DAppDate.Style.Add("background-color", "greenyellow")
                DAppDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DAppDateRqd", "DAppDate", "異常：需輸入預定完成日")
                BDate3.Visible = False

            Case 2  '修改
                DAppDate.Visible = True
                DAppDate.Style.Add("background-color", "yellow")
                DAppDate.Attributes.Add("readonly", "true")
                BDate3.Visible = False

            Case Else   '隱藏
                DAppDate.Visible = False
                BDate3.Visible = False

        End Select

        If pPost = "New" Then DAppDate.Value = CDate(NowDateTime).ToString("yyyy/MM/dd")








        '擔當
        Select Case FindFieldInf("Appper")
            Case 0  '顯示
                DAppper.BackColor = Color.LightGray
                DAppper.Visible = True
                DAppper.ReadOnly = True
            Case 1  '修改+檢查
                DAppper.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppperRqd", "DAppper", "異常：需輸入擔當")
                DAppper.Visible = True
            Case 2  '修改
                DAppper.BackColor = Color.Yellow
                DAppper.Visible = True
            Case Else   '隱藏
                DAppper.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Appper", "ZZZZZZ")


        '依賴者
        Select Case FindFieldInf("Clinter")
            Case 0  '顯示
                DClinter.BackColor = Color.LightGray
                DClinter.Visible = True
                DClinter.ReadOnly = True
                BClienter.Visible = False
            Case 1  '修改+檢查
                DClinter.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DClinterRqd", "DClinter", "異常：需輸入依賴者")
                DClinter.Visible = True
                '    DClinter.ReadOnly = True
                BClienter.Visible = False
            Case 2  '修改
                DClinter.BackColor = Color.Yellow
                DClinter.Visible = True
                DClinter.ReadOnly = True
                BClienter.Visible = False
            Case Else   '隱藏
                DClinter.Visible = False
                BClienter.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Clinter", "ZZZZZZ")


        '部門
        Select Case FindFieldInf("DivisionCode")
            Case 0  '顯示
                DDivisionCode.BackColor = Color.LightGray
                DDivisionCode.Visible = True

            Case 1  '修改+檢查
                DDivisionCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionCodeRqd", "DDivisionCode", "異常：需輸入類別1")
                DDivisionCode.Visible = True
            Case 2  '修改
                DDivisionCode.BackColor = Color.Yellow
                DDivisionCode.Visible = True
            Case Else   '隱藏
                DDivisionCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DivisionCode", "ZZZZZZ")




        '部門
        '  Select Case FindFieldInf("Division")
        '      Case 0  '顯示
        ' DDivision.BackColor = Color.LightGray
        ' DDivision.Visible = True
        ' DDivision.ReadOnly = True
        '     Case 1  '修改+檢查
        ' DDivision.BackColor = Color.GreenYellow
        'ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
        'DDivision.Visible = True
        'DDivision.ReadOnly = True
        '    Case 2  '修改
        'DDivision.BackColor = Color.Yellow
        'DDivision.Visible = True
        '    Case Else   '隱藏
        'DDivision.Visible = False
        'End Select
        'If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")



        '類別1
        Select Case FindFieldInf("Type1")
            Case 0  '顯示
                DType1.BackColor = Color.LightGray
                DType1.Visible = True

            Case 1  '修改+檢查
                DType1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DType1Rqd", "DType1", "異常：需輸入類別1")
                DType1.Visible = True
            Case 2  '修改
                DType1.BackColor = Color.Yellow
                DType1.Visible = True
            Case Else   '隱藏
                DType1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Type1", "ZZZZZZ")


        '類別2
        Select Case FindFieldInf("Type2")
            Case 0  '顯示
                DType2.BackColor = Color.LightGray
                DType2.Visible = True

            Case 1  '修改+檢查
                DType2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DType2Rqd", "DType2", "異常：需輸入類別2")
                DType2.Visible = True
            Case 2  '修改
                DType2.BackColor = Color.Yellow
                DType2.Visible = True
            Case Else   '隱藏
                DType2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Type2", "ZZZZZZ")


        '類別2
        Select Case FindFieldInf("Engine")
            Case 0  '顯示
                DEngine.BackColor = Color.LightGray
                DEngine.Visible = True


            Case 1  '修改+檢查
                DEngine.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DEngineRqd", "DEngine", "異常：需輸入工程")
                DEngine.Visible = True

            Case 2  '修改
                DEngine.BackColor = Color.Yellow
                DEngine.Visible = True

            Case Else   '隱藏
                DEngine.Visible = False


        End Select
        If pPost = "New" Then SetFieldData("Engine", "ZZZZZZ")



        '圖號
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.LightGray
                DMapNo.Visible = True
                DMapNo.ReadOnly = True
            Case 1  '修改+檢查
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "異常：需輸入圖號")
                DMapNo.Visible = True
            Case 2  '修改
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
            Case Else   '隱藏
                DMapNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MapNo", "ZZZZZZ")

        '數量
        Select Case FindFieldInf("Qty")
            Case 0  '顯示
                DQty.BackColor = Color.LightGray
                DQty.Visible = True
                DQty.ReadOnly = True
            Case 1  '修改+檢查
                DQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQtyRqd", "DQty", "異常：需輸入數量")
                DQty.Visible = True
            Case 2  '修改
                DQty.BackColor = Color.Yellow
                DQty.Visible = True
            Case Else   '隱藏
                DQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Qty", "ZZZZZZ")

        '數量
        Select Case FindFieldInf("AQty")
            Case 0  '顯示
                DAQty.BackColor = Color.LightGray
                DAQty.Visible = True
                DAQty.ReadOnly = True
            Case 1  '修改+檢查
                DAQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAQtyRqd", "DAQty", "異常：需輸入數量")
                DAQty.Visible = True
            Case 2  '修改
                DAQty.BackColor = Color.Yellow
                DAQty.Visible = True
            Case Else   '隱藏
                DAQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AQty", "ZZZZZZ")




        'Code
        Select Case FindFieldInf("Code")
            Case 0  '顯示
                DCode.BackColor = Color.LightGray
                DCode.Visible = True
                DCode.ReadOnly = True
            Case 1  '修改+檢查
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "異常：需輸入數量")
                DCode.Visible = True
            Case 2  '修改
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '隱藏
                DCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Code", "ZZZZZZ")





        '品名
        Select Case FindFieldInf("Weight")
            Case 0  '顯示
                DWeight.BackColor = Color.LightGray
                DWeight.Visible = True
                DWeight.ReadOnly = True
            Case 1  '修改+檢查
                DWeight.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWeightRqd", "DWeight", "異常：需輸入品名")
                DWeight.Visible = True
            Case 2  '修改
                DWeight.BackColor = Color.Yellow
                DWeight.Visible = True
            Case Else   '隱藏
                DWeight.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Weight", "ZZZZZZ")



        '材料
        Select Case FindFieldInf("Material")
            Case 0  '顯示
                DMaterial.BackColor = Color.LightGray
                DMaterial.Visible = True
                DMaterial.ReadOnly = True
            Case 1  '修改+檢查
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "異常：需輸入品名")
                DMaterial.Visible = True
            Case 2  '修改
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '隱藏
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")




        '材料
        Select Case FindFieldInf("DevNo")
            Case 0  '顯示
                DDevNo.BackColor = Color.LightGray
                DDevNo.Visible = True
                DDevNo.ReadOnly = True
            Case 1  '修改+檢查
                DDevNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevNoRqd", "DDevNo", "異常：需輸入品名")
                DDevNo.Visible = True
            Case 2  '修改
                DDevNo.BackColor = Color.Yellow
                DDevNo.Visible = True
            Case Else   '隱藏
                DDevNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DevNo", "ZZZZZZ")


        '材料
        Select Case FindFieldInf("Manufout")
            Case 0  '顯示
                DManufout.BackColor = Color.LightGray
                DManufout.Visible = True
                DManufout.ReadOnly = True
            Case 1  '修改+檢查
                DManufout.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufoutRqd", "DManufout", "異常：需輸入品名")
                DManufout.Visible = True
            Case 2  '修改
                DManufout.BackColor = Color.Yellow
                DManufout.Visible = True
            Case Else   '隱藏
                DManufout.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Manufout", "ZZZZZZ")



        '預定完成日
        Select Case FindFieldInf("FinishDate")
            Case 0  '顯示
                DFinishDate.Visible = True
                DFinishDate.Style.Add("background-color", "lightgrey")
                DFinishDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = True


            Case 1  '修改+檢查
                DFinishDate.Visible = True
                DFinishDate.Style.Add("background-color", "greenyellow")
                DFinishDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DFinishDateRqd", "DFinishDate", "異常：需輸入預定完成日")
                BDate1.Disabled = False

            Case 2  '修改
                DFinishDate.Visible = True
                DFinishDate.Style.Add("background-color", "yellow")
                DFinishDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = False

            Case Else   '隱藏
                DFinishDate.Visible = False
                BDate1.Disabled = True

        End Select
        If pPost = "New" Then DFinishDate.Value = ""



        '完成日
        Select Case FindFieldInf("AFinishDate")
            Case 0  '顯示
                DAFinishDate.Visible = True
                DAFinishDate.Style.Add("background-color", "lightgrey")
                DAFinishDate.Attributes.Add("readonly", "true")
                BDate2.Visible = False

            Case 1  '修改+檢查
                DAFinishDate.Visible = True
                DAFinishDate.Style.Add("background-color", "greenyellow")
                DAFinishDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DAFinishDateRqd", "DAFinishDate", "異常：需輸入預定完成日")
                BDate2.Visible = True


            Case 2  '修改
                DAFinishDate.Visible = True
                DAFinishDate.Style.Add("background-color", "yellow")
                DAFinishDate.Attributes.Add("readonly", "true")
                BDate2.Visible = True


            Case Else   '隱藏
                DAFinishDate.Visible = False
                BDate2.Visible = False

        End Select

        If pPost = "New" Then DAFinishDate.Value = CDate(NowDateTime).ToString("yyyy/MM/dd")




        '簡圖
        Select Case FindFieldInf("MapFile")
            Case 0  '顯示
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "異常：需輸入簡圖")
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMapFile.Visible = False
        End Select
        LMapFile.Visible = False
        LMapFile1.Visible = False

        '工程1
        Select Case FindFieldInf("Engine1")
            Case 0  '顯示
                DEngine1.BackColor = Color.LightGray
                DEngine1.Visible = True
                DEngine1.ReadOnly = True


            Case 1  '修改+檢查
                DEngine1.BackColor = Color.GreenYellow
                '   ShowRequiredFieldValidator("DEngine1Rqd", "DEngine1", "異常：需輸入品名")
                DEngine1.Visible = True

            Case 2  '修改
                DEngine1.BackColor = Color.Yellow
                DEngine1.Visible = True
                DEngine1.ReadOnly = True

            Case Else   '隱藏
                DEngine1.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("Engine1", "ZZZZZZ")

        '工程2
        Select Case FindFieldInf("Engine2")
            Case 0  '顯示
                DEngine2.BackColor = Color.LightGray
                DEngine2.Visible = True
                DEngine2.ReadOnly = True
            Case 1  '修改+檢查
                DEngine2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine2Rqd", "DEngine2", "異常：需輸入品名")
                DEngine2.Visible = True
            Case 2  '修改
                DEngine2.BackColor = Color.Yellow
                DEngine2.Visible = True
                DEngine2.ReadOnly = True
            Case Else   '隱藏
                DEngine2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine2", "ZZZZZZ")

        '工程3
        Select Case FindFieldInf("Engine3")
            Case 0  '顯示
                DEngine3.BackColor = Color.LightGray
                DEngine3.Visible = True
                DEngine3.ReadOnly = True
            Case 1  '修改+檢查
                DEngine3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine3Rqd", "DEngine3", "異常：需輸入品名")
                DEngine3.Visible = True
            Case 2  '修改
                DEngine3.BackColor = Color.Yellow
                DEngine3.Visible = True
                DEngine3.ReadOnly = True
            Case Else   '隱藏
                DEngine3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine3", "ZZZZZZ")


        '工程3
        Select Case FindFieldInf("Engine4")
            Case 0  '顯示
                DEngine4.BackColor = Color.LightGray
                DEngine4.Visible = True
                DEngine4.ReadOnly = True
            Case 1  '修改+檢查
                DEngine4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine4Rqd", "DEngine4", "異常：需輸入品名")
                DEngine4.Visible = True
            Case 2  '修改
                DEngine4.BackColor = Color.Yellow
                DEngine4.Visible = True
                DEngine4.ReadOnly = True
            Case Else   '隱藏
                DEngine4.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine4", "ZZZZZZ")



        '工程3
        Select Case FindFieldInf("Engine5")
            Case 0  '顯示
                DEngine5.BackColor = Color.LightGray
                DEngine5.Visible = True
                DEngine5.ReadOnly = True
            Case 1  '修改+檢查
                DEngine5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine5Rqd", "DEngine5", "異常：需輸入品名")
                DEngine5.Visible = True
            Case 2  '修改
                DEngine5.BackColor = Color.Yellow
                DEngine5.Visible = True
                DEngine5.ReadOnly = True
            Case Else   '隱藏
                DEngine5.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine5", "ZZZZZZ")


        '工程3
        Select Case FindFieldInf("Engine6")
            Case 0  '顯示
                DEngine6.BackColor = Color.LightGray
                DEngine6.Visible = True
                DEngine6.ReadOnly = True
            Case 1  '修改+檢查
                DEngine6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine6Rqd", "DEngine6", "異常：需輸入品名")
                DEngine6.Visible = True
            Case 2  '修改
                DEngine6.BackColor = Color.Yellow
                DEngine6.Visible = True
                DEngine6.ReadOnly = True
            Case Else   '隱藏
                DEngine6.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine6", "ZZZZZZ")


        '工程3
        Select Case FindFieldInf("Engine7")
            Case 0  '顯示
                DEngine7.BackColor = Color.LightGray
                DEngine7.Visible = True
                DEngine7.ReadOnly = True
            Case 1  '修改+檢查
                DEngine7.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine7Rqd", "DEngine7", "異常：需輸入品名")
                DEngine7.Visible = True
            Case 2  '修改
                DEngine7.BackColor = Color.Yellow
                DEngine7.Visible = True
                DEngine7.ReadOnly = True
            Case Else   '隱藏
                DEngine7.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine7", "ZZZZZZ")



        '工程3
        Select Case FindFieldInf("Engine8")
            Case 0  '顯示
                DEngine8.BackColor = Color.LightGray
                DEngine8.Visible = True
                DEngine8.ReadOnly = True
            Case 1  '修改+檢查
                DEngine8.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine8Rqd", "DEngine8", "異常：需輸入品名")
                DEngine8.Visible = True
            Case 2  '修改
                DEngine8.BackColor = Color.Yellow
                DEngine8.Visible = True
                DEngine8.ReadOnly = True
            Case Else   '隱藏
                DEngine8.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine8", "ZZZZZZ")


        '工程3
        Select Case FindFieldInf("Engine9")
            Case 0  '顯示
                DEngine9.BackColor = Color.LightGray
                DEngine9.Visible = True
                DEngine9.ReadOnly = True
            Case 1  '修改+檢查
                DEngine9.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine9Rqd", "DEngine9", "異常：需輸入品名")
                DEngine9.Visible = True
            Case 2  '修改
                DEngine9.BackColor = Color.Yellow
                DEngine9.Visible = True
                DEngine9.ReadOnly = True
            Case Else   '隱藏
                DEngine9.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine9", "ZZZZZZ")

        '工程3
        Select Case FindFieldInf("Engine10")
            Case 0  '顯示
                DEngine10.BackColor = Color.LightGray
                DEngine10.Visible = True
                DEngine10.ReadOnly = True
            Case 1  '修改+檢查
                DEngine10.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine10Rqd", "DEngine10", "異常：需輸入品名")
                DEngine10.Visible = True
            Case 2  '修改
                DEngine10.BackColor = Color.Yellow
                DEngine10.Visible = True
                DEngine10.ReadOnly = True
            Case Else   '隱藏
                DEngine10.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine10", "ZZZZZZ")

        '工程3
        Select Case FindFieldInf("Engine11")
            Case 0  '顯示
                DEngine11.BackColor = Color.LightGray
                DEngine11.Visible = True
                DEngine11.ReadOnly = True
            Case 1  '修改+檢查
                DEngine11.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine11Rqd", "DEngine11", "異常：需輸入品名")
                DEngine11.Visible = True
            Case 2  '修改
                DEngine11.BackColor = Color.Yellow
                DEngine11.Visible = True
                DEngine11.ReadOnly = True
            Case Else   '隱藏
                DEngine11.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine11", "ZZZZZZ")

        '工程3
        Select Case FindFieldInf("Engine12")
            Case 0  '顯示
                DEngine12.BackColor = Color.LightGray
                DEngine12.Visible = True
                DEngine12.ReadOnly = True
            Case 1  '修改+檢查
                DEngine12.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine12Rqd", "DEngine12", "異常：需輸入品名")
                DEngine12.Visible = True
            Case 2  '修改
                DEngine12.BackColor = Color.Yellow
                DEngine12.Visible = True
                DEngine12.ReadOnly = True
            Case Else   '隱藏
                DEngine12.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine12", "ZZZZZZ")

        '工程3
        Select Case FindFieldInf("Engine13")
            Case 0  '顯示
                DEngine13.BackColor = Color.LightGray
                DEngine13.Visible = True
                DEngine13.ReadOnly = True
            Case 1  '修改+檢查
                DEngine13.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine13Rqd", "DEngine13", "異常：需輸入品名")
                DEngine13.Visible = True
            Case 2  '修改
                DEngine13.BackColor = Color.Yellow
                DEngine13.Visible = True
                DEngine13.ReadOnly = True
            Case Else   '隱藏
                DEngine13.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine13", "ZZZZZZ")



        '工程3
        Select Case FindFieldInf("Engine14")
            Case 0  '顯示
                DEngine14.BackColor = Color.LightGray
                DEngine14.Visible = True
                DEngine14.ReadOnly = True
            Case 1  '修改+檢查
                DEngine14.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine14Rqd", "DEngine14", "異常：需輸入品名")
                DEngine14.Visible = True
            Case 2  '修改
                DEngine14.BackColor = Color.Yellow
                DEngine14.Visible = True
                DEngine14.ReadOnly = True
            Case Else   '隱藏
                DEngine14.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine14", "ZZZZZZ")


    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.None              ' 因在頁面上加入ValidationSummary , 故驗證控制項統一顯示
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer

        idx = FindFieldInf(pFieldName)


        '擔當者及部門 

        '擔當者及部門 

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DAppper.Text = DBUser.Rows(0).Item("Username")

        'DNo.Text = SetNo(wFormSno)




        If pFieldName = "DivisionCode" Then
            DDivisionCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then


                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivisionCode.Items.Add(ListItem1)
                End If
            Else
                SQL = " select  distinct dep_code+'-'+dep_name Data  from M_emp"
                SQL = SQL + " where com_code  in ('01','60','65','70','71')"
                SQL = SQL + " union all"
                SQL = SQL + " select top 1 '1010010-工廠共通' Data from M_emp"
                SQL = SQL + " union all"
                SQL = SQL + " select top 1 '1010331-工機(SPD)' Data from M_emp "
                SQL = SQL + " union all"
                SQL = SQL + " select top 1 '1010331-工機(不良)' Data from M_emp "
                SQL = SQL + " order by data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DDivisionCode.Items.Add("")
                'DDivisionCode.Items.Add("1010010-工廠共通")
                'DDivisionCode.Items.Add("1010331-工機(SPD)")
                'DDivisionCode.Items.Add("1010331-工機(不良)")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then
                        ListItem1.Selected = True
                    End If

                    DDivisionCode.Items.Add(ListItem1)
                Next
            End If

        End If





        If pFieldName = "Type1" Then
            DType1.Items.Clear()

            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DType1.Items.Add(ListItem1)

                End If
            Else
                SQL = "Select * From M_Referp Where Cat='4001' and dkey='TYPE1' Order by DKey, Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DType1.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DType1.Items.Add(ListItem1)
                Next
            End If

        End If

        If pFieldName = "Type2" Then
            DType2.Items.Clear()

            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DType2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='4001' and dkey='TYPE2' Order by DKey, Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DType2.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DType2.Items.Add(ListItem1)
                Next
            End If



        End If

        If pFieldName = "Engine" Then
            DEngine.Items.Clear()

            If pName <> "ZZZZZZ" Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = pName
                ListItem1.Value = pName
                DEngine.Items.Add(ListItem1)

            Else
                SQL = "Select distinct substring(dkey,14,len(dkey)-1)data  From M_Referp Where Cat='4001' and substring(dkey,1,12) =  'EngineSelect' Order by Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DEngine.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEngine.Items.Add(ListItem1)
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
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("DKey")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If

        '延遲理由
        If pFieldName = "Reason" Then
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DReason.Text = pName
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                For i = 0 To DBAdapter1.Rows.Count - 1
                    DReason.Text = DBAdapter1.Rows(i).Item("Data")
                Next
            End If
        End If



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


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick

        AQty = DAQty.Text
        DevNo = DDevNo.Text
        Manufout = DManufout.Text

        If Request.Cookies("RunBOK").Value = True Then

            If (wStep = 10) Or (wStep = 20) Or (wStep = 30) Or (wStep = 40) Then
                If OK() = True Then
                    DisabledButton()   '停止Button運作
                    FlowControl("OK", 0, "1")
                End If
            Else
                DisabledButton()   '停止Button運作
                FlowControl("OK", 0, "1")
            End If
        End If


    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG1按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then


            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG2按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then

            If wStep = 10 Then
                If OK() = True Then
                    DisabledButton()   '停止Button運作
                    FlowControl("NG2", 2, "3")

                End If
            Else
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")

            End If
        End If
    End Sub

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
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '取得此關卡各按鈕的Data Status



        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("004001", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '樣品難易度

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
                Dim i, RtnCode As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
                Dim SQL As String

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        '取得表單流水號
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else

                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo,wUserID, wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, wUserID, wApplyID)
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
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,wUserID, pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, wUserID, pRunNextStep)
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

                    Dim wAllocateID As String = ""
                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, wUserID, wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

                    If wStep = 15 Then

                    End If

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
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo,wUserID, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
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
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, wUserID, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo,wUserID, wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, wUserID, wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(wUserID, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(wUserID, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        'Modify-End
                    End If
                End If
                '儲存表單資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If  'pSeqno <> 0
                    Else    '判斷是否起單
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
                            oCommon.Send(wUserID, pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        oCommon.Send("mcs006a", wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
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
                '--郵件傳送---------
                oCommon.SendMail()
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                 "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "材質為其他時需填寫說明,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9070 Then Message = "需填入工程待處理部門及工程待處理者.!"
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
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                   System.Configuration.ConfigurationManager.AppSettings("MPMProcessesFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String


        SQl = "Insert into F_MPMProcessesSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Appdate,APPPER,Clinter,DivisionCode,"                  '1~5
        SQl = SQl + "Division,Type1,Type2,MapNo,Qty,AQty,Code,"                  '6~10
        SQl = SQl + "Weight,Material,DevNo,Manufout, FinishDate,"
        If wStep = 40 Then
            SQl = SQl + "AFinishDate, "
        End If

        SQl = SQl + " MapFile, Aranger, Arang, "                  '11~13"
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '004001', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號
        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        SQl = SQl + " '" + DAppDate.Value + "', "                '日期
        SQl = SQl + " '" + DAppper.Text + "', "   '部門
        SQl = SQl + " '" + DClinter.Text + "', "     '擔當
        SQl = SQl + " '" + Mid(DDivisionCode.SelectedValue, 1, 7) + "', "  'buyer
        SQl = SQl + " '" + Mid(DDivisionCode.SelectedValue, 9, Len(DDivisionCode.SelectedValue) - 1) + "',"
        SQl = SQl + " '" + DType1.SelectedValue + "', "  'buyer
        SQl = SQl + " '" + DType2.SelectedValue + "', "  'buyer
        SQl = SQl + " '" + DMapNo.Text + "', "  'buyer
        SQl = SQl + " '" + DQty.Text + "', "  'buyer
        SQl = SQl + " '" + DAQty.Text + "', "  'buyer
        SQl = SQl + " '" + DCode.Text + "', "  'buyer
        SQl = SQl + " '" + DWeight.Text + "', "  'buyer
        SQl = SQl + " '" + DMaterial.Text + "', "  'buyer
        SQl = SQl + " '" + DDevNo.Text + "', "  'buyer
        SQl = SQl + " '" + DManufout.Text + "', "  'buyer


        If DFinishDate.Value = "" Then
            SQl = SQl + " '', "  'buyer
        Else
            SQl = SQl + " '" + DFinishDate.Value + "', "  'buyer
        End If

        If wStep = 40 Then
            If DAFinishDate.Value = "" Then
                SQl = SQl + " '', "  'buyer
            Else
                SQl = SQl + " '" + DAFinishDate.Value + "', "  'buyer
            End If


        End If



        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)

                DMapFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '品質依賴書 
        SQl = SQl + " '', "  'Aranger
        SQl = SQl + " 0," 'Arange
        SQl = SQl + " '" + wUserID + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
        SQl = SQl + " ) "

        uDataBase.ExecuteNonQuery(SQl)

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                   System.Configuration.ConfigurationManager.AppSettings("MPMProcessesFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim SQL1 As String = ""
        Dim FileName As String = ""
        Dim i As Integer


        If wStep = 15 Then
            SQL1 = "Insert into F_MPMProcessesSheetDT "
            SQL1 = SQL1 + "( "
            SQL1 = SQL1 + "FormSno,Seqno,No,Engine)"
            SQL1 = SQL1 + "Values('" + Str(wFormSno) + "','1','" + YKK.ReplaceString(DNo.Text) + "','CAM')"
            uDataBase.ExecuteNonQuery(SQL1)

        End If


        '加入工程細項
        Dim EngineStr As String

        If wStep = 10 Or wStep = 20 Then
            'SQL1 = "Delete F_MPMProcessesSheetDT "
            'SQL1 = SQL1 + " where Formsno = '" + Str(wFormSno) + "'"
            'SQL1 = SQL1 + " and recdate is   null"
            'uDataBase.ExecuteNonQuery(SQL1)

            '20221104 Jessica 如果第一筆有刷記錄就從2開始
            Dim seqno As Integer = 0
            SQL1 = " select isnull(max(seqno),'0')Data from  F_MPMProcessesSheetDT "
            SQL1 = SQL1 + " where Formsno = '" + Str(wFormSno) + "'"

            Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL1)
            If dt.Rows.Count > 0 Then
                seqno = dt.Rows(0)("Data")
            End If





            For i = 1 + seqno To 14
                EngineStr = "DEngine" + Trim(Str(i))
                Dim DText As TextBox = Me.FindControl(EngineStr)
                If DText.Text <> "" Then

                    SQL1 = "Insert into F_MPMProcessesSheetDT "
                    SQL1 = SQL1 + "( "
                    SQL1 = SQL1 + "FormSno,Seqno,No,Engine)"
                    SQL1 = SQL1 + "Values('" + Str(wFormSno) + "',N'" + Str(i) + "','" + YKK.ReplaceString(DNo.Text) + "','" + DText.Text + "')"
                    uDataBase.ExecuteNonQuery(SQL1)

                    ' 更新安排工程人員()

                End If

            Next

        End If


  




        SQL1 = " Update F_MPMProcessesSheet"
        SQL1 = SQL1 + " Set "
        If pFun <> "SAVE" Then
            SQL1 = SQL1 + " Sts = '" & pSts & "',"
            SQL1 = SQL1 + " CompletedTime = '" & NowDateTime & "',"
        End If

        SQL1 = SQL1 + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL1 = SQL1 + " AppDate = '" & DAppDate.Value & "',"
        SQL1 = SQL1 + " Appper = '" & DAppper.Text & "',"
        SQL1 = SQL1 + " Clinter = '" & DClinter.Text & "',"
        SQL1 = SQL1 + " DivisionCode = '" & Mid(DDivisionCode.SelectedValue, 1, 7) & "',"
        SQL1 = SQL1 + " Division = '" & Mid(DDivisionCode.SelectedValue, 9, Len(DDivisionCode.SelectedValue) - 1) & "',"
        SQL1 = SQL1 + " Type1 = '" & DType1.SelectedValue & "',"
        SQL1 = SQL1 + " Type2 = '" & DType2.SelectedValue & "',"
        SQL1 = SQL1 + " MapNo = '" & DMapNo.Text & "',"
        SQL1 = SQL1 + " Qty = '" & Trim(DQty.Text) & "',"
        SQL1 = SQL1 + " AQty = '" & AQty & "',"
        SQL1 = SQL1 + " Code = '" & DCode.Text & "',"
        SQL1 = SQL1 + " Weight = '" & DWeight.Text & "',"
        SQL1 = SQL1 + " Material = '" & DMaterial.Text & "',"
        SQL1 = SQL1 + " DevNo = '" & DevNo & "',"
        SQL1 = SQL1 + " Manufout = '" & Manufout & "',"

        SQL1 = SQL1 + " FinishDate = '" & DFinishDate.Value & "',"
        If wStep = 40 Then
            SQL1 = SQL1 + " AFinishDate = '" & DAFinishDate.Value & "',"
        End If


       
        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                DMapFile.PostedFile.SaveAs(Path + FileName)
                SQL1 = SQL1 + " MapFile= N'" + YKK.ReplaceString(FileName) + "',"           '品質依賴書
            Else
                FileName = ""
            End If
        End If

        If wStep = 10 Or wStep = 20 Or wStep = 30 Then
            SQL1 = SQL1 + " Aranger = '" & Request.QueryString("pUserID") & "',"
        End If

        SQL1 = SQL1 + " Arang = 1" & ","
        SQL1 = SQL1 + " ModifyUser = '" & wUserID & "',"
        SQL1 = SQL1 + " ModifyTime = '" & NowDateTime & "' "

        SQL1 = SQL1 + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL1)






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
            SQl = SQl + " ModifyUser = '" & wUserID & "',"
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
            SQl = SQl + " ModifyUser = '" & wUserID & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        uDataBase.ExecuteNonQuery(SQl)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String


        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQl)

        If DBAdapter1.Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + wUserID + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & wUserID & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    uDataBase.ExecuteNonQuery(SQl)
                End If
            End If
        End If

    End Sub

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
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()

        If wStep = 40 Then
            Top = Top - 250
        End If

        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 250 & "px"
            DDecideDesc.Style("top") = Top - 250 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If

        If wStep = 40 Then
            Top = Top - 100
        End If

        BSAVE.Style("top") = Top + 10 & "px"
        BNG1.Style("top") = Top + 10 & "px"
        BNG2.Style("top") = Top + 10 & "px"
        BOK.Style("top") = Top + 10 & "px"

        Top += 48

        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
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
        '  Str2 = CStr(DateTime.Now.Month)  '月
        '  For i = Str2.Length To 1
        ' Str2 = "0" + Str2
        ' Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '日
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = CStr(DateTime.Now.Year)  '年
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 10 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    Protected Sub DEngine_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DEngine.SelectedIndexChanged
        Dim i As Integer
        i = 30
        BEngineAdd.Visible = True
        BEngineDel.Visible = True

        '顯示工程內容

        If DEngine.SelectedValue = "單獨選取" Then
            '展開工程選項 
            'Engine-A 單獨選取

            DPanel.Visible = True
            CheckBoxList1.RepeatColumns = 8
            CheckBoxList1.RepeatDirection = RepeatDirection.Horizontal
            CheckBoxList1.Visible = True
            CheckBoxList1.Items.Clear()
            RadioButtonList1.Visible = False
            DPanel.BackColor = Color.Aquamarine


            Dim SQL As String

            SQL = "select * from M_referp "
            SQL = SQL + "where cat = '4001'"
            SQL = SQL + "and dkey ='EngineSelect-單獨選取'"
            SQL = SQL + " Order by Unique_ID"

            Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

            For Each dtr As Data.DataRow In dt.Rows
                Dim EngineA As New ListItem(dtr("Data"), dtr("Data"))
                'New ListItem(dtr("RUserName"), dtr("RUserID"))
                CheckBoxList1.Items.Add(EngineA)
            Next
            i = 90
            '顯示高度
            DPanel.Height = i
            Top = 1000
        ElseIf DEngine.SelectedValue <> "" Then

            '模具區、零件區
            DPanel.Visible = True
            CheckBoxList1.Visible = False
            RadioButtonList1.Visible = True
            If DEngine.SelectedValue = "零件區" Then
                DPanel.BackColor = Color.Violet
                Top = 1200
            Else
                DPanel.BackColor = Color.LightGreen
                Top = 1100
            End If



            '組合工程字串
            Dim SQL As String


            SQL = "select * from M_referp "
            SQL = SQL + "where cat = '4001'"
            SQL = SQL + "and dkey ='EngineSelect-" + DEngine.SelectedValue + "'"
            SQL = SQL + " Order by Unique_ID"
            Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
            RadioButtonList1.Items.Clear()
            For Each dtr As Data.DataRow In dt.Rows
                Dim EngineStr As String = ""
                SQL = "select * from M_referp "
                SQL = SQL + "where cat = '4001'"
                SQL = SQL + "and  dkey = '" + dtr("Data") + "'"
                SQL = SQL + " Order by Unique_ID"
                Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)
                For Each dtr1 As Data.DataRow In dt1.Rows
                    If EngineStr = "" Then
                        EngineStr = dtr1("Data")
                    Else
                        EngineStr = EngineStr + "," + dtr1("Data")
                    End If

                Next
                Dim EngineB As New ListItem(dtr("Data") + "【" + EngineStr + "】", dtr("Data"))
                'New ListItem(dtr("RUserName"), dtr("RUserID"))
                RadioButtonList1.Items.Add(EngineB)
                i = i + 30
            Next

            DPanel.Height = i
        Else
            BEngineAdd.Visible = False
            BEngineDel.Visible = False
            DPanel.Visible = False
        End If


        SetControlPosition()
        ShowFormData()
    End Sub


    Protected Sub BEngineAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEngineAdd.Click
        Dim i, j, k As Integer
        Dim CheckStr, EngineStr As String

        '填入工程
        If DEngine1.Text = "CAM" Then
            k = 2
        Else
            k = 1
        End If

        For i = k To 14


            EngineStr = "DEngine" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = ""

        Next


        If DEngine.SelectedValue = "單獨選取" Then
            j = k

            For i = 0 To (CheckBoxList1.Items.Count - 1)
                If (CheckBoxList1.Items(i).Selected) Then
                    CheckStr = Trim(CheckBoxList1.Items(i).Text)

                    EngineStr = "DEngine" + Trim(Str(j))
                    Dim DText As TextBox = Me.FindControl(EngineStr)
                    'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)
                    If DText IsNot Nothing Then
                        DText.Text = CheckStr
                    End If
                    j = j + 1
                End If

            Next

            Top = 1000
        Else
            CheckStr = ""
            For i = 0 To (RadioButtonList1.Items.Count - 1)
                If RadioButtonList1.Items(i).Selected Then
                    CheckStr = Trim(RadioButtonList1.Items(i).Value)
                End If
            Next

            Dim SQL As String
            SQL = "select Data from M_referp "
            SQL = SQL + "where cat = '4001'"
            SQL = SQL + " and dkey = '" + CheckStr + "'"
            SQL = SQL + " Order by Unique_ID"
            Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)
            k = k - 1
            For i = 0 To dt1.Rows.Count - 1

                CheckStr = dt1.Rows(i).Item("Data")
                EngineStr = "DEngine" + Trim(Str(i + 1 + k))
                Dim DText As TextBox = Me.FindControl(EngineStr)
                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)
                If DText IsNot Nothing Then
                    DText.Text = CheckStr
                End If
            Next

            If DEngine.SelectedValue = "零件區" Then

                Top = 1200
            Else

                Top = 1100
            End If


        End If

        SetControlPosition()
        ShowFormData()
    End Sub

    Protected Sub BEngineDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEngineDel.Click
        '清除資料
        Dim i As Integer
        Dim EngineStr As String
        For i = 1 To 14
            EngineStr = "DEngine" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = ""

        Next

        For i = 0 To (CheckBoxList1.Items.Count - 1)
            CheckBoxList1.Items(i).Selected = False

        Next


        Top = 1200
        SetControlPosition()
        ShowFormData()
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function OK() As Boolean
        Top = 1200
        SetControlPosition()
        ShowFormData()

        Dim isOK As Boolean = True
        Dim Errcode As Integer = 0



        Dim Message As String = ""


        Dim i, j As Integer
        Dim EngineStr As String

        If wStep = 10 Or wStep = 20 Or wStep = 30 Then
            j = 0
            For i = 1 To 14
                EngineStr = "DEngine" + Trim(Str(i))
                Dim DText As TextBox = Me.FindControl(EngineStr)
                If DText.Text <> "" Then
                    j = j + 1
                End If


            Next

            If j = 0 Then
                isOK = False
                Message = "異常：需輸入工程!"
            End If

        End If
     
        If wStep = 40 Then '檢查是否可結案
            Dim SQL As String
            SQL = " select * from  F_MPMProcessesSheetdt "
            SQL = SQL + "  where  no ='" + DNo.Text + "'"
            SQL = SQL + " and engine <> '' "
            SQL = SQL + " and (starter ='' "
            SQL = SQL + " or ender  ='') "
            SQL = SQL + " and delmark = 0 "
            Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)
            If dt1.Rows.Count > 0 Then
                j = 1
            End If

            'SQL = " select * from  F_MPMProcessesSheetdt "
            'SQL = SQL + "  where  no ='" + DNo.Text + "'"
            'SQL = SQL + "and ender  <>'' "
            'SQL = SQL + "and delmark = 0 "
            'Dim dt3 As Data.DataTable = uDataBase.GetDataTable(SQL)
            'If dt3.Rows.Count > 0 Then
            '    j = 0
            'End If


            SQL = " select * from  F_MPMProcessesSheetdt "
            SQL = SQL + "  where  no ='" + DNo.Text + "'"
            SQL = SQL + "and ( AlwaysStoper <>'')"
            SQL = SQL + "and delmark = 0 "
            Dim dt2 As Data.DataTable = uDataBase.GetDataTable(SQL)
            If dt2.Rows.Count > 0 Then
                j = 0
            End If



            If j = 1 Then
                isOK = False
                Message = "異常：尚有工程未完成，不可結案!"
            End If

        End If

    

     
        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If


        Return isOK

      
    End Function

 
 

    Protected Sub DAppper_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAppper.TextChanged

    End Sub

    Protected Sub DNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.TextChanged

    End Sub
End Class

