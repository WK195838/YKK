Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet01_UAKIPLINGDT
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
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
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
    Dim isOK As Boolean = True
    Dim Message As String = ""
    Dim LastStep As String






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
        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼

        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)

        wUserID = Request.QueryString("pUserID")

        'wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "DTMW_NewColorSheet01_03CFP12.aspx"
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

        SQL = "Select * From F_NewColorUAKIPLINGDT "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DNO1.Text = DBAdapter1.Rows(0).Item("No1")

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")

            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))

            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

            SetFieldData("DTSheet", DBAdapter1.Rows(0).Item("DTSheet"))




        End If


        '帶入最新屢歷
        Dim sqlhistory As String
        sqlhistory = " SELECT"
        sqlhistory = sqlhistory + " a.No  As Field1,"
        sqlhistory = sqlhistory + " case b.Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2,"
        sqlhistory = sqlhistory + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3,"
        sqlhistory = sqlhistory + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate,"
        sqlhistory = sqlhistory + " a.FormName as Field4,"
        sqlhistory = sqlhistory + " a.Division As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"
        sqlhistory = sqlhistory + " YKKColorType As Field9,YKKColorCode as Field10,SLDColor As Field11,VFColor As Field12,NewOldColor,"
        sqlhistory = sqlhistory + " '....' as WorkFlow, ViewURL,"
        sqlhistory = sqlhistory + " 'http://10.245.1.10/WorkFlow/BefOPList.aspx?' +"
        sqlhistory = sqlhistory + " 'pFormNo='   + a.FormNo +"
        sqlhistory = sqlhistory + " '&pFormSno=' + str(a.FormSno,Len(a.FormSno)) +"
        sqlhistory = sqlhistory + " '&pStep='    + str(a.Step,Len(a.Step)) +"
        sqlhistory = sqlhistory + " '&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) +"
        sqlhistory = sqlhistory + " '&pApplyID=' + a.ApplyID"
        sqlhistory = sqlhistory + " As OPURL, "
        sqlhistory = sqlhistory + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,"
        sqlhistory = sqlhistory + " customerColorCode,overSeaYkkCode,pantonecode,substring(stepnamedesc,7,len(stepnamedesc )-1)stepnamedesc,a.FormSno "
        sqlhistory = sqlhistory + " from V_WaitHandle_01 a,V_NewColor b"
        sqlhistory = sqlhistory + " Where a.formno=b.formno and a.formsno =b.formsno and active  = '1' "
        sqlhistory = sqlhistory + " and a.no ='" & DNO1.Text & "'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(sqlhistory)

        If DBAdapter3.Rows.Count > 0 Then

           
            HyperLink1.NavigateUrl = DBAdapter3.Rows(0).Item("OPURL")  '請假證明
            HyperLink1.Visible = True
            MNOSts.Text = DBAdapter3.Rows(0).Item("stepnamedesc")
            MNOSts.Visible = True
            DOFormSno.Text = DBAdapter3.Rows(0).Item("FormSno")
            HyperLink2.NavigateUrl = "DTMW_NewColorSheet02_UAKIPLING.aspx?pFormNo=005011" & "&pFormSno=" & DOFormSno.Text
            HyperLink2.Visible = True
            Dim i As Integer

            MNOSts.Text = ""

            For i = 0 To DBAdapter3.Rows.Count - 1
                If MNOSts.Text = "" Then
                    MNOSts.Text = DBAdapter3.Rows(i).Item("stepnamedesc")
                Else
                    MNOSts.Text = MNOSts.Text + "," + DBAdapter3.Rows(i).Item("stepnamedesc")
                End If

            Next


        End If





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
                    Top = 580
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    If DBAdapter2.Rows(0).Item("Flowtype") = "1" Then
                        Top = 580
                    Else
                        Top = 580
                    End If

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
            Top = 580
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

        BCopySheet.Attributes.Add("onClick", "CopyNewColor('" + wFormNo + "')") '找buyer


        BYKKColorCode.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCode')") '找buyer
        Me.DColorSystem.Attributes.CssStyle.Add("text-transform", "uppercase")

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
                    Top = 520
                Else
                    If DDelay.Visible = True Then
                        Top = 520
                    Else
                        Top = 520
                    End If
                End If
            End If
        Else
            Top = 520

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


        'Dim SQL As String


        'MNo
        Select Case FindFieldInf("NO1")
            Case 0  '顯示
                DNO1.BackColor = Color.LightGray
                DNO1.ReadOnly = True
                DNO1.Visible = True
                BCopySheet.Visible = False
            Case 1  '修改+檢查
                DNO1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNO1Rqd", "DNO1", "異常：需輸入Ｎｏ")
                DNO1.Visible = True
                BCopySheet.Visible = True
            Case 2  '修改
                DNO1.BackColor = Color.Yellow
                DNO1.Visible = True
                BCopySheet.Visible = True
            Case Else   '隱藏
                DNO1.Visible = False
                BCopySheet.Visible = False
        End Select
        If pPost = "New" Then DNO1.Text = ""





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
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True

            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DDate.Visible = True

            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.Visible = True

            Case Else   '隱藏
                DDate.Visible = False

        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '修改+檢查
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入部門")
                DDepoName.Visible = True
            Case 2  '修改
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DepoName", "ZZZZZZ")


        'Name
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.ReadOnly = True
                DName.Visible = True
            Case 1  '修改+檢查
                DName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
                DName.Visible = True
            Case 2  '修改
                DName.BackColor = Color.Yellow
                DName.Visible = True
            Case Else   '隱藏
                DName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Name", "ZZZZZZ")





        'SLDColor
        Select Case FindFieldInf("SLDColor")
            Case 0  '顯示
                DSLDColor.BackColor = Color.LightGray
                DSLDColor.ReadOnly = True
                DSLDColor.Visible = True

            Case 1  '修改+檢查
                DSLDColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSLDColorRqd", "DSLDColor", "異常：需輸入確認拉頭兼用色")
                DSLDColor.Visible = True

            Case 2  '修改
                DSLDColor.BackColor = Color.Yellow
                DSLDColor.Visible = True
            Case Else   '隱藏
                DSLDColor.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("SLDColor", "ZZZZZZ")

        'VFColor 
        Select Case FindFieldInf("VFColor")
            Case 0  '顯示
                DVFColor.BackColor = Color.LightGray
                DVFColor.ReadOnly = True
                DVFColor.Visible = True
            Case 1  '修改+檢查
                DVFColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColorRqd", "DVFColor", "異常：需輸入確認VF兼用色")
                DVFColor.Visible = True

            Case 2  '修改
                DVFColor.BackColor = Color.Yellow
                DVFColor.Visible = True
            Case Else   '隱藏
                DVFColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor", "ZZZZZZ")







        'YKK色別
        Select Case FindFieldInf("YKKColorType")
            Case 0  '顯示
                DYKKColorType.BackColor = Color.LightGray
                DYKKColorType.Visible = True

            Case 1  '修改+檢查
                DYKKColorType.BackColor = Color.GreenYellow
                '      ShowRequiredFieldValidator("DYKKColorTypeRqd", "DYKKColorType", "異常：需輸YKK色別")
                DYKKColorType.Visible = True
            Case 2  '修改
                DYKKColorType.BackColor = Color.Yellow
                DYKKColorType.Visible = True
            Case Else   '隱藏
                DYKKColorType.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("YKKColorType", "ZZZZZZ")




        'YKKColorCode
        Select Case FindFieldInf("YKKColorCode")
            Case 0  '顯示
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "lightgrey")
                DYKKColorCode.Attributes.Add("readonly", "true")
                BYKKColorCode.Visible = False

            Case 1  '修改+檢查
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "greenyellow")
                '  DYKKColorCode.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorCodeRqd", "DYKKColorCode", "異常：需輸入YKK色號")
                BYKKColorCode.Visible = True

            Case 2  '修改
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "yellow")
                DYKKColorCode.Attributes.Add("readonly", "true")
                BYKKColorCode.Visible = True
            Case Else   '隱藏
                DYKKColorCode.Visible = False
                BYKKColorCode.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCode.Value = ""



        'PFBWire


        'PFBWire 
        Select Case FindFieldInf("PFBWire")
            Case 0  '顯示
                DPFBWire.BackColor = Color.LightGray
                DPFBWire.ReadOnly = True
                DPFBWire.Visible = True
            Case 1  '修改+檢查
                DPFBWire.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFBWireRqd", "DPFBWire", "異常：需輸入確認VF兼用色")
                DPFBWire.Visible = True

            Case 2  '修改
                DPFBWire.BackColor = Color.Yellow
                DPFBWire.Visible = True
            Case Else   '隱藏
                DPFBWire.Visible = False
        End Select
        If pPost = "New" Then DPFBWire.Text = ""



        'DPFOpenParts 
        Select Case FindFieldInf("PFOpenParts")
            Case 0  '顯示
                DPFOpenParts.BackColor = Color.LightGray
                DPFOpenParts.ReadOnly = True
                DPFOpenParts.Visible = True
            Case 1  '修改+檢查
                DPFOpenParts.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFOpenPartsRqd", "DPFOpenParts", "異常：需輸入確認VF兼用色")
                DPFOpenParts.Visible = True

            Case 2  '修改
                DPFOpenParts.BackColor = Color.Yellow
                DPFOpenParts.Visible = True
            Case Else   '隱藏
                DPFOpenParts.Visible = False
        End Select
        If pPost = "New" Then DPFOpenParts.Text = ""




        'YKK色別
        Select Case FindFieldInf("ColorSystem")
            Case 0  '顯示
                DColorSystem.BackColor = Color.LightGray
                DColorSystem.Visible = True

            Case 1  '修改+檢查
                DColorSystem.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DColorSystemRqd", "DColorSystem", "異常：需輸色系")
                DColorSystem.Visible = True
            Case 2  '修改
                DColorSystem.BackColor = Color.Yellow
                DColorSystem.Visible = True
            Case Else   '隱藏
                DColorSystem.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorSystem", "ZZZZZZ")




        '新舊色
        Select Case FindFieldInf("NewOldColor")
            Case 0  '顯示
                DNewOldColor.BackColor = Color.LightGray
                DNewOldColor.Visible = True

            Case 1  '修改+檢查
                DNewOldColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNewOldColorRqd", "DNewOldColor", "異常：需輸新舊色")
                DNewOldColor.Visible = True
            Case 2  '修改
                DNewOldColor.BackColor = Color.Yellow
                DNewOldColor.Visible = True
            Case Else   '隱藏
                DNewOldColor.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("NewOldColor", "ZZZZZZ")


        'ReRegister
        Select Case FindFieldInf("ReRegister")
            Case 0  '顯示
                DReRegister.BackColor = Color.LightGray
                DReRegister.Visible = True
            Case 1  '修改+檢查
                DReRegister.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReRegisterRqd", "DReRegister", "異常：需輸入繼續申請")
                DReRegister.Visible = True
            Case 2  '修改
                DReRegister.BackColor = Color.Yellow
                DReRegister.Visible = True
            Case Else   '隱藏
                DReRegister.Visible = False
                DReRegister.Checked = False
        End Select

        '核可單種類
        Select Case FindFieldInf("DTSheet")
            Case 0  '顯示
                DDTSheet.BackColor = Color.LightGray
                DDTSheet.Visible = True

            Case 1  '修改+檢查
                DDTSheet.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDTSheetRqd", "DDTSheet", "異常：需輸核可單種類")
                DDTSheet.Visible = True
            Case 2  '修改
                DDTSheet.BackColor = Color.Yellow
                DDTSheet.Visible = True
            Case Else   '隱藏
                DDTSheet.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("DTSheet", "ZZZZZZ")


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

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DDepoName.Text = DBUser.Rows(0).Item("Divname")
        DName.Text = DBUser.Rows(0).Item("Username")



        'YKK色別
        If pFieldName = "YKKColorType" Then
            DYKKColorType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYKKColorType.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'YKKColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DYKKColorType.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYKKColorType.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        '新色舊色
        If pFieldName = "NewOldColor" Then
            DNewOldColor.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNewOldColor.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'NewOldColor'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DNewOldColor.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNewOldColor.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '核可單種類
        If pFieldName = "DTSheet" Then
            DDTSheet.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDTSheet.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DTSheet'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDTSheet.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDTSheet.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
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
        If Request.Cookies("RunBSAVE").Value = True And OK() = True Then

            DisabledButton()   '停止Button運作
            FlowControl("OK", 3, "1")

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True And OK() = True Then

            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")

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
            If DNO1.Text <> "" Then
                ErrCode = oCommon.CommissionNo("005012", wFormSno, wStep, DNO1.Text) '表單號碼, 表單流水號, 工程, 委託書No
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

                    If wStep = 10 Then

                        If DDTSheet.SelectedValue = "布帶牙齒" Then
                            pAction = 2
                        Else

                            If DNewOldColor.SelectedValue = "新色" Then
                                If DDTSheet.SelectedValue = "布帶" Then
                                    pAction = 0
                                Else
                                    pAction = 3
                                End If
                            Else
                                If DYKKColorType.SelectedValue = "螢光" Then
                                    pAction = 1  '151
                                Else
                                    pAction = 2
                                End If
                            End If

                        End If




                    End If

                    If wStep = 20 Then
                        If DYKKColorType.SelectedValue = "螢光" Then
                            pAction = 0
                        Else
                            pAction = 1
                        End If
                    End If

                    If wStep = 152 Then
                        If DDTSheet.SelectedValue = "布帶牙齒" Then
                            pAction = 1
                        ElseIf DDTSheet.SelectedValue = "拉頭" Then
                            pAction = 2
                        Else
                            pAction = 0
                        End If
                    End If


                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, wUserID, wApplyID, wAgentID, wAllocateID, MultiJob, _
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
                        oCommon.Send(wUserID, wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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


                    LastStep = wStep  '記錄上一個工程
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

                Dim URL As String

                ' MODIFY-START BY JOY 2015/4/7 (徐慧芬)
                ' If (LastStep = 20 Or LastStep = 520 Or LastStep = 25) And pFun = "OK" Then '新色確認完成表按OK才進來
                '增加連續起單 Modify 20150419
                If wbFormSno = 0 And wbStep < 3 And DReRegister.Checked = True Then    '判斷是否起單

                    uJavaScript.PopMsg(Me, "已成功送出您所填寫之申請單，請繼續申請！")
                    EnabledButton()                             '起動Button運作


                    DNO1.Text = ""
                    DNo.Text = ""
                    HyperLink1.Text = ""
                    MNOSts.Text = ""

                    wFormSno = wbFormSno
                    wStep = wbStep


                Else



                    URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                "&pUserID=" & wUserID & "&pApplyID=" & wApplyID

                    Response.Redirect(URL)
                End If '

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

        Dim SQl As String
        SQl = "Insert into F_NewColorUAKIPLINGDT "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, DepoName,Date,Name,colorsystem,"
        SQl = SQl + "PFBwire,PFOpenParts,YKKColorType,YKKColorCode,"
        SQl = SQl + "SLDColor,VFCOLOR,NewOldColor,oFormNo,oFormSno,DTSheet,MNOSts,No1,"
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '005012', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號
        SQl = SQl + " N'" + DNo.Text + "', "                '部門2
        SQl = SQl + " N'" + DDepoName.Text + "', "                '日期3
        SQl = SQl + " N'" + DDate.Text + "', "                '姓名4
        SQl = SQl + " N'" + DName.Text + "', "                '部門2
        SQl = SQl + " N'" + DColorSystem.Text + "', "                '日期3
        SQl = SQl + " N'" + DPFBWire.Text + "', "
        SQl = SQl + " N'" + DPFOpenParts.Text + "', "
        SQl = SQl + " N'" + DYKKColorType.Text + "', "
        SQl = SQl + " N'" + DYKKColorCode.Value + "', "
        SQl = SQl + " N'" + DSLDColor.Text.ToUpper + "', "
        SQl = SQl + " N'" + DVFColor.Text.ToUpper + "', "
        SQl = SQl + " N'" + DNewOldColor.Text.ToUpper + "', "
        SQl = SQl + " '005011',"
        SQl = SQl + " N'" + DOFormSno.Text + "', "
        SQl = SQl + " N'" + DDTSheet.Text + "', "
        SQl = SQl + " N'" + MNOSts.Text + "', "
        SQl = SQl + " N'" + DNO1.Text + "', "
        SQl = SQl + " '" + wUserID + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)

        SQl = " Update F_NewColorUAKIPLING"
        SQl = SQl + " Set "
        SQl = SQl + " DTNO =1 "
        SQl = SQl + " ,OFormNo ='005012' "
        SQl = SQl + " ,OFormsNo ='" + CStr(NewFormSno) + "'"
        SQl = SQl + " ,DTSheet = '" + DDTSheet.Text + "'"
        SQl = SQl + " where formsno ='" + DOFormSno.Text + "'"
        uDataBase.ExecuteNonQuery(SQl)

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim SQL As String = ""

        SQL = " Update F_NewColorUAKIPLINGDT"
        SQL = SQL + " Set "
        If pFun <> "SAVE" Then
            SQL = SQL + " Sts = '" & pSts & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQL = SQL + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = N'" & DDate.Text & "',"
        SQL = SQL + " DepoName = N'" & DDepoName.Text & "',"
        SQL = SQL + " Name = N'" & DName.Text & "',"
        SQL = SQL + " ColorSystem = N'" & DColorSystem.Text.ToUpper & "',"
        SQL = SQL + " PFBWire = N'" & DPFBWire.Text & "',"
        SQL = SQL + " PFOpenParts = N'" & DPFOpenParts.Text & "',"
        SQL = SQL + " YKKColorType = N'" & DYKKColorType.Text & "',"
        SQL = SQL + " YKKColorCode = N'" & DYKKColorCode.Value & "',"
        SQL = SQL + " SLDColor = N'" & DSLDColor.Text.ToUpper & "',"
        SQL = SQL + " VFColor = N'" & DVFColor.Text.ToUpper & "',"
        SQL = SQL + " NewOldColor = N'" & DNewOldColor.Text & "',"
        SQL = SQL + " oFormNo= '005011',"
        SQL = SQL + " oFormSno = N'" & DOFormSno.Text & "',"
        SQL = SQL + " DTSheet = N'" & DDTSheet.Text & "',"
        SQL = SQL + " MNoSts = N'" & MNOSts.Text & "',"
        SQL = SQL + " No1 = N'" & DNO1.Text & "',"
        SQL = SQL + " ModifyUser = '" & wUserID & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL)




        SQL = " Update F_NewColorUAKIPLING"
        SQL = SQL + " Set "
        SQL = SQL + " DTNO =1 "
        SQL = SQL + " ,DTSheet = '" + DDTSheet.Text + "'"
        SQL = SQL + " ,NewoldColor='" + DNewOldColor.SelectedValue + "'"

        If DDTSheet.SelectedValue = "布帶" Then
            SQL = SQL + " ,YKKColorCode = '" + DYKKColorCode.Value + "'"
        ElseIf DDTSheet.SelectedValue = "布帶牙齒" Then
            SQL = SQL + " ,YKKColorCode = '" + DYKKColorCode.Value + "'"
            SQL = SQL + " ,YKKColorCodeVF = '" + DYKKColorCode.Value + "'"

        ElseIf DDTSheet.SelectedValue = "牙齒" Then
            SQL = SQL + " ,YKKColorCodeVF = '" + DYKKColorCode.Value + "'"
        Else
            SQL = SQL + " ,YKKColorCodeSLD = '" + DYKKColorCode.Value + "'"
        End If

        SQL = SQL + " where formsno ='" + DOFormSno.Text + "'"
        uDataBase.ExecuteNonQuery(SQL)


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
                If DNO1.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
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



        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 30 & "px"
            DDecideDesc.Style("top") = Top - 30 + 6 & "px"
            Top = Top + 70
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"


        Top = Top + 30
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
        Dim SQL As String = ""

        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '日
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = "K" + CStr(DateTime.Now.Year) + Str2
        '年
        'Set單號
        '當月份筆數有幾筆  20150414 Modify by Jessica
        SQL = " select   isnull(max(convert(int,substring(no,8,4))),0) cun from  F_NewColorUAKIPLINGDT"
        SQL = SQL + " where  left(convert(char(10),date ,111),7) = left(convert(char(10), getdate(),111),7)"
        Dim dt1 As DataTable = uDataBase.GetDataTable(SQL)
        If dt1.Rows.Count > 0 Then
            Str1 = CStr(CInt(dt1.Rows(0).Item("cun")) + 1)
        End If

        ' Str1 = CStr(Seq)

        For i = Str1.Length To 4 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function OK() As Boolean
        Top = 600
        SetControlPosition()
        '  ShowFormData()
        Dim SQL As String
        Dim DOUBLENo As Integer = 0

        Dim Errcode As Integer = 0





        'Dim Q As Integer = 0
        ''jessica 20150326
        'If wStep = 10 Then
        '    '檢查是否有Q 
        '    Q = InStr(1, DYKKColorCode.Value, "Q", 1)
        '    If DYKKColorType.SelectedValue = "螢光" Then


        '        SQL = " select * from   m_referp"
        '        SQL = SQL + " where dkey  = 'LightCode'"
        '        SQL = SQL + " and cat = 5001"
        '        SQL = SQL + " and data  ='" + DYKKColorCode.Value + "'"

        '        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        '        If DBAdapter1.Rows.Count > 0 Then

        '        Else
        '            If Q = 0 Then
        '                Message = Message + "\n" + "只有Q字樣選項才能選螢光!"
        '                DYKKColorCode.Value = ""
        '                isOK = False
        '            End If
        '        End If
        '    Else
        '        If Q > 0 Then
        '            Message = Message + "\n" + "有Q字樣選項只能選螢光!"
        '            DYKKColorCode.Value = ""
        '            isOK = False
        '        End If
        '    End If


        'End If






        If Not isOK Then

            uJavaScript.PopMsg(Me, Message)



        End If




        Return isOK


    End Function


    Sub MsgBox(ByVal text As String)
        'Dim scriptstr As String
        'scriptstr = "<script language=javascript>" + Chr(10) _
        '+ "confirm(""" + text + """)" + Chr(10) _
        '+ "</script>"
        'Response.Write(scriptstr)
        'Response.Write("<script language='javascript'>if(confirm(""" + text + """)==true)function1();else function2();</script>")
        'Response.Write("<script language=javascript>if (confirm('你確定要繼續輸出文字？')==false) {return fale;};</script>")
    End Sub

    Sub CopyNo()
        'COPY 表單


        Dim NO1, sql, sqlhistory As String
        NO1 = DNO1.Text
        If NO1 <> "" Then
            sql = "Select * From V_NewColor "
            sql = sql & " Where no  =  '" & NO1 & "'"

            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            If DBAdapter1.Rows.Count > 0 Then
                DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")
                DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
                DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")
                DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
                DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

                DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
                DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

                SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '類別2
                SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '類別2

            End If

            sqlhistory = " SELECT"
            sqlhistory = sqlhistory + " a.No  As Field1,"
            sqlhistory = sqlhistory + " case b.Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2,"
            sqlhistory = sqlhistory + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3,"
            sqlhistory = sqlhistory + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate,"
            sqlhistory = sqlhistory + " a.FormName as Field4,"
            sqlhistory = sqlhistory + " a.Division As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"
            sqlhistory = sqlhistory + " YKKColorType As Field9,YKKColorCode as Field10,SLDColor As Field11,VFColor As Field12,NewOldColor,"
            sqlhistory = sqlhistory + " '....' as WorkFlow, ViewURL,"
            sqlhistory = sqlhistory + " 'http://10.245.1.10/WorkFlow/BefOPList.aspx?' +"
            sqlhistory = sqlhistory + " 'pFormNo='   + a.FormNo +"
            sqlhistory = sqlhistory + " '&pFormSno=' + str(a.FormSno,Len(a.FormSno)) +"
            sqlhistory = sqlhistory + " '&pStep='    + str(a.Step,Len(a.Step)) +"
            sqlhistory = sqlhistory + " '&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) +"
            sqlhistory = sqlhistory + " '&pApplyID=' + a.ApplyID"
            sqlhistory = sqlhistory + " As OPURL, "
            sqlhistory = sqlhistory + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,"
            sqlhistory = sqlhistory + " customerColorCode,overSeaYkkCode,pantonecode,substring(stepnamedesc,7,len(stepnamedesc )-1)stepnamedesc,a.FormSno "
            sqlhistory = sqlhistory + " from V_WaitHandle_01 a,V_NewColor b"
            sqlhistory = sqlhistory + " Where a.formno=b.formno and a.formsno =b.formsno and active  = '1' "
            sqlhistory = sqlhistory + " and a.no ='" & NO1 & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sqlhistory)
            Dim i As Integer
            If DBAdapter2.Rows.Count > 0 Then
                HyperLink1.NavigateUrl = DBAdapter2.Rows(0).Item("OPURL")  ' 
                HyperLink1.Visible = True


                MNOSts.Text = ""

                For i = 0 To DBAdapter2.Rows.Count - 1
                    If MNOSts.Text = "" Then
                        MNOSts.Text = DBAdapter2.Rows(i).Item("stepnamedesc")
                    Else
                        MNOSts.Text = MNOSts.Text + "," + DBAdapter2.Rows(i).Item("stepnamedesc")
                    End If

                Next



                'MNOSts.Text = DBAdapter2.Rows(0).Item("stepnamedesc")
                MNOSts.Visible = True
                DOFormSno.Text = DBAdapter2.Rows(0).Item("FormSno")
                DDTSheet.Text = ""
                ' DYKKColorType.SelectedValue = ""
                ' DNewOldColor.SelectedValue = ""
            End If

        End If
    End Sub


    Protected Sub DNO1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNO1.TextChanged
        CopyNo()
    End Sub


End Class

