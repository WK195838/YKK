Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDCommissionSheet_01
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
    Dim flag As Integer
    Dim PAction1 As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        ' TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置
        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                SetControlPosition()    ' 設定控制項位置
            End If
            SetPopupFunction()      '設定彈出視窗事件

        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text, Request.QueryString("pUserID")) ' 設定預設的簽核者
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
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        Response.Cookies("PGM").Value = "SBDCommissionSheet_01.aspx"
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
        Dim SQL As String

        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDCommissionFilePath")



        SQL = "Select * From F_SBDCommissionSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            SetFieldData("Buyer", DBAdapter1.Rows(0).Item("Buyer"))    'Buyer
            DVendor.Text = DBAdapter1.Rows(0).Item("Vendor")              'Vendor
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DBackground.Text = DBAdapter1.Rows(0).Item("Background")              'Backround
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              'Code
            DMapDate.Value = DBAdapter1.Rows(0).Item("MapDate")              'MapDate
            DSampleDate.Value = DBAdapter1.Rows(0).Item("SampleDate")              'SampleDate
            SetFieldData("Light", DBAdapter1.Rows(0).Item("Light")) 'Light    

            DHalfFinishNo.Text = DBAdapter1.Rows(0).Item("HalfFinishNo")              'HalfFinishNot
            SetFieldData("Material", DBAdapter1.Rows(0).Item("Material"))    'Material
            SetFieldData("MaterialDetail", DBAdapter1.Rows(0).Item("MaterialDetail"))    'MaterialDetail
            DMaterialDetail_1.Text = DBAdapter1.Rows(0).Item("MaterialDetail_1")              'MaterialDetail_1



            If DBAdapter1.Rows(0).Item("RefMapFile") <> "" Then
                LRefMapFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("RefMapFile")  'LCertifcateFile
                LRefMapFile.Visible = True
            Else
                LRefMapFile.Visible = False
            End If




            SetFieldData("Sample", DBAdapter1.Rows(0).Item("Sample"))    'Sample


            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")              'Remark
            DMapNo.Text = DBAdapter1.Rows(0).Item("MapNo")              'MapNo
            SetFieldData("MakeMap", "")    'MakeMap
            DMakeMapU.Text = DBAdapter1.Rows(0).Item("MakeMap")              'MapNo
            SetFieldData("Level", DBAdapter1.Rows(0).Item("Level"))    'Level
            'LMapFile

            If Trim(DBAdapter1.Rows(0).Item("MapFile")) <> "" Then
                LMapFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  'LMapFile
                LMapFile.Visible = True
            Else
                LMapFile.Visible = False
            End If


            'SetFieldData("2-Makemap1", DBAdapter1.Rows(0).Item("MakeMap1"))    'MakeMap
            'SetFieldData("2-Makemap2", DBAdapter1.Rows(0).Item("MakeMap2"))    'MakeMap
            'SetFieldData("2-Makemap3", DBAdapter1.Rows(0).Item("MakeMap3"))    'MakeMap
            'SetFieldData("2-Makemap4", DBAdapter1.Rows(0).Item("MakeMap4"))    'MakeMap
            'SetFieldData("2-Makemap5", DBAdapter1.Rows(0).Item("MakeMap5"))    'MakeMap
            'SetFieldData("2-Makemap6", DBAdapter1.Rows(0).Item("MakeMap6"))    'MakeMap
            DMakeMap1.Text = DBAdapter1.Rows(0).Item("MakeMap1")
            DMakeMap2.Text = DBAdapter1.Rows(0).Item("MakeMap2")
            DMakeMap3.Text = DBAdapter1.Rows(0).Item("MakeMap3")
            DMakeMap4.Text = DBAdapter1.Rows(0).Item("MakeMap4")
            DMakeMap5.Text = DBAdapter1.Rows(0).Item("MakeMap5")
            DMakeMap6.Text = DBAdapter1.Rows(0).Item("MakeMap6")




            '  Dim a As String
            'a = DBAdapter1.Rows(0).Item("MakeMap1")
            ' DMakemap1.Text = DBAdapter1.Rows(0).Item("MakeMap1")


            SetFieldData("Reason1", DBAdapter1.Rows(0).Item("Reason1"))    'Reason

            DContent1.Text = DBAdapter1.Rows(0).Item("Content1")              'DContent1

            If DBAdapter1.Rows(0).Item("ContentFile1") <> "" Then
                LContentFile1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile1")  'LContentFile1 
                LContentFile1.Visible = True
            Else
                LContentFile1.Visible = False
            End If
            DContent2.Text = DBAdapter1.Rows(0).Item("Content2")              'DContent2

            If DBAdapter1.Rows(0).Item("ContentFile2") <> "" Then
                LContentFile2.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile2")  'LContentFile2 
                LContentFile2.Visible = True
            Else
                LContentFile2.Visible = False
            End If

            DContent3.Text = DBAdapter1.Rows(0).Item("Content3")              'DContent3

            If DBAdapter1.Rows(0).Item("ContentFile3") <> "" Then
                LContentFile3.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile3")  'LContentFile3 
                LContentFile3.Visible = True
            Else
                LContentFile3.Visible = False
            End If

            DContent4.Text = DBAdapter1.Rows(0).Item("Content4")              'DContent4

            If DBAdapter1.Rows(0).Item("ContentFile4") <> "" Then
                LContentFile4.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile4")  'LContentFile4 
                LContentFile4.Visible = True
            Else
                LContentFile4.Visible = False
            End If

            DContent5.Text = DBAdapter1.Rows(0).Item("Content5")              'DContent5

            If DBAdapter1.Rows(0).Item("ContentFile5") <> "" Then
                LContentFile5.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile5")  'LContentFile5 
                LContentFile5.Visible = True
            Else
                LContentFile5.Visible = False
            End If

            DContent6.Text = DBAdapter1.Rows(0).Item("Content6")              'DContent6

            If DBAdapter1.Rows(0).Item("ContentFile6") <> "" Then
                LContentFile6.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile6")  'LContentFile6 
                LContentFile6.Visible = True
            Else
                LContentFile6.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map1") <> "" Then
                LMap1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map1")  'LMap1
                LMap1.Visible = True
            Else
                LMap1.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map2") <> "" Then
                LMap2.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map2")  'LMap1
                LMap2.Visible = True
            Else
                LMap2.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map3") <> "" Then
                LMap3.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map3")  'LMap1
                LMap3.Visible = True
            Else
                LMap3.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map4") <> "" Then
                LMap4.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map4")  'LMap1
                LMap4.Visible = True
            Else
                LMap4.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map5") <> "" Then
                LMap5.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map5")  'LMap1
                LMap5.Visible = True
            Else
                LMap5.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map6") <> "" Then
                LMap6.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map6")  'LMap1
                LMap6.Visible = True
            Else
                LMap6.Visible = False
            End If








            SetFieldData("Supplier", DBAdapter1.Rows(0).Item("Supplier"))    'Suppiler

            DHalfFinishOrderNo.Text = DBAdapter1.Rows(0).Item("HalfFinishOrderNo")              'HalfFinishOrderNo
            If Mid(DBAdapter1.Rows(0).Item("HalfFinishDate").ToString, 1, 4) = "1900" Then
                DHalfFinishDate.Value = ""
            Else
                DHalfFinishDate.Value = DBAdapter1.Rows(0).Item("HalfFinishDate")              'HalfFinishDdate
            End If

            DMold.Text = DBAdapter1.Rows(0).Item("Mold")              'Mold

            DMoldPoint.Text = DBAdapter1.Rows(0).Item("MoldPoint")              'MoldPoint

            If DBAdapter1.Rows(0).Item("Surfcolor") <> "" Then
                DSurfcolor.Value = DBAdapter1.Rows(0).Item("Surfcolor")              'Surfcolor
                LSurfColor.Visible = True
                LSurfColor.Text = "________"
                LSurfColor.NavigateUrl = "SBDSurfaceSheet_02.aspx?pFormNo=" + RTrim(CStr(DBAdapter1.Rows(0).Item("SurfFormNo"))) + "&pFormSno=" + RTrim(CStr(DBAdapter1.Rows(0).Item("SurfFormSno")))

            End If
            'DSurfcolor.Value = DBAdapter1.Rows(0).Item("Surfcolor")              'Surfcolor
            DSurfcolor1.Text = DBAdapter1.Rows(0).Item("Surfcolor1")              'Surfcolor
            'DSampleQty.Text = DBAdapter1.Rows(0).Item("SampleQty")


            SetFieldData("SampleQty", DBAdapter1.Rows(0).Item("SampleQty"))   'Suppiler

            If DBAdapter1.Rows(0).Item("QCReqFile") <> "" Then
                LQCReqFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QCReqFile")  'LQCReqFile
                LQCReqFile.Visible = True
            Else
                LQCReqFile.Visible = False
            End If

            DFQA.Text = DBAdapter1.Rows(0).Item("FQA")              'FQA
            DQARemark.Text = DBAdapter1.Rows(0).Item("QARemark")              'QARemark



            If DBAdapter1.Rows(0).Item("ForcastFile") <> "" Then
                LForcastFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ForcastFile")  ' LForcastFile
                LForcastFile.Visible = True
            Else
                LForcastFile.Visible = False
            End If



            If DBAdapter1.Rows(0).Item("QAFile") <> "" Then
                LQAFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QAFile")  ' LQAFile
                LQAFile.Visible = True
            Else
                LQAFile.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("AuthorizeFile") <> "" Then
                LAuthorizeFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("AuthorizeFile")  ' LAuthorizeFile
                LAuthorizeFile.Visible = True
            Else
                LAuthorizeFile.Visible = False
            End If


            If DBAdapter1.Rows(0).Item("SampleFile") <> "" Then
                LSampleFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("SampleFile")  ' LSampleFile
                LSampleFile.Visible = True
            Else
                LSampleFile.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("ContactFile") <> "" Then
                LContactFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContactFile")  'LContactFile
                LContactFile.Visible = True
            Else
                LContactFile.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("RefFile") <> "" Then
                LRefFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("RefFile")  'LContactFile
                LRefFile.Visible = True
            Else
                LRefFile.Visible = False
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
                    Top = 1200
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 1500
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
            Top = 1400
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
        BDate1.Attributes.Add("onClick", "calendarPicker('DMapDate')")
        BDate2.Attributes.Add("onClick", "calendarPicker('DSampleDate')")
        BDate3.Attributes.Add("onClick", "calendarPicker('DHalfFinishDate')")
        BColor.Attributes.Add("onClick", "openNoPicker('DSurfcolor')")
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
                    Top = 1600
                Else
                    If DDelay.Visible = True Then
                        Top = 1600
                    Else
                        Top = 1600
                    End If
                End If
            End If
        Else
            Top = 1600
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
                DAppDate.BackColor = Color.LightGray
                DAppDate.ReadOnly = True
                DAppDate.Visible = True

            Case 1  '修改+檢查
                DAppDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppDateRqd", "DDate", "異常：需輸入日期")
                DAppDate.Visible = True

            Case 2  '修改
                DAppDate.BackColor = Color.Yellow
                DAppDate.Visible = True

            Case Else   '隱藏
                DAppDate.Visible = False

        End Select
        If pPost = "New" Then DAppDate.Text = Now.ToString("yyyy/MM/dd") '現在日時


        'buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.Visible = True

            Case 1  '修改+檢查
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
            Case 2  '修改
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '隱藏
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")

        'Vendor
        Select Case FindFieldInf("Vendor")
            Case 0  '顯示
                DVendor.BackColor = Color.LightGray
                DVendor.ReadOnly = False
                DVendor.Visible = True
            Case 1  '修改+檢查
                DVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVendorRqd", "DVendor", "異常：需輸入Vendor")
                DVendor.Visible = True
            Case 2  '修改
                DVendor.BackColor = Color.Yellow
                DVendor.Visible = True
            Case Else   '隱藏
                DVendor.Visible = False
        End Select
        If pPost = "New" Then DVendor.Text = ""

        '修改部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.ReadOnly = True
                DDivision.Visible = True
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")

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

        '開發背景
        Select Case FindFieldInf("Background")
            Case 0  '顯示
                DBackground.BackColor = Color.LightGray
                DBackground.Visible = True
                DBackground.ReadOnly = True
            Case 1  '修改+檢查
                DBackground.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBackgroundtRqd", "DBackground", "異常：需輸入開發背景")
                DBackground.Visible = True
            Case 2  '修改
                DBackground.BackColor = Color.Yellow
                DBackground.Visible = True
            Case Else   '隱藏
                DBackground.Visible = False
        End Select
        If pPost = "New" Then DBackground.Text = ""

        'Code
        Select Case FindFieldInf("Code")
            Case 0  '顯示
                DCode.BackColor = Color.LightGray
                DCode.ReadOnly = True
                DCode.Visible = True
            Case 1  '修改+檢查
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "異常：需輸入Code")
                DCode.Visible = True
            Case 2  '修改
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '隱藏
                DCode.Visible = False
        End Select
        If pPost = "New" Then DCode.Text = ""

        '圖面希望交期
        Select Case FindFieldInf("MapDate")
            Case 0  '顯示
                DMapDate.Visible = True
                DMapDate.Style.Add("background-color", "lightgrey")
                DMapDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = True
                BDate1.Disabled = True
            Case 1  '修改+檢查
                DMapDate.Visible = True
                DMapDate.Style.Add("background-color", "greenyellow")
                DMapDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DMapDateRqd", "DMapDate", "異常：需輸入圖面希望交期")
                BDate1.Disabled = False
                BDate1.Disabled = False
            Case 2  '修改
                DMapDate.Visible = True
                DMapDate.Style.Add("background-color", "yellow")
                DMapDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = False
                BDate1.Disabled = False
            Case Else   '隱藏
                DMapDate.Visible = False
                BDate1.Disabled = True
                BDate1.Disabled = True
        End Select
        If pPost = "New" Then DMapDate.Value = ""

        '樣品希望交期

        Select Case FindFieldInf("SampleDate")
            Case 0  '顯示
                DSampleDate.Visible = True
                DSampleDate.Style.Add("background-color", "lightgrey")
                DSampleDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = True
                BDate2.Disabled = True
            Case 1  '修改+檢查
                DSampleDate.Visible = True
                DSampleDate.Style.Add("background-color", "greenyellow")
                DSampleDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DSampledateRqd", "Dsampledate", "異常：需輸入樣品希望交期")
                BDate2.Disabled = False
                BDate2.Disabled = False
            Case 2  '修改
                DSampleDate.Visible = True
                DSampleDate.Style.Add("background-color", "yellow")
                DSampleDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = False
                BDate2.Disabled = False
            Case Else   '隱藏
                DSampleDate.Visible = False
                BDate2.Disabled = True
                BDate2.Disabled = True
        End Select
        If pPost = "New" Then DSampleDate.Value = ""

        '光造型
        Select Case FindFieldInf("Light")
            Case 0  '顯示
                DLight.BackColor = Color.LightGray
                DLight.Visible = True

            Case 1  '修改+檢查
                DLight.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLightRqd", "DLight", "異常：需輸入光造型")
                DLight.Visible = True
            Case 2  '修改
                DLight.BackColor = Color.Yellow
                DLight.Visible = True
            Case Else   '隱藏
                DLight.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Light", "ZZZZZZ")

        '半成品料號
        Select Case FindFieldInf("HalffinishNo")
            Case 0  '顯示
                DHalfFinishNo.BackColor = Color.LightGray
                DHalfFinishNo.Visible = True
                DHalfFinishNo.ReadOnly = True
            Case 1  '修改+檢查
                DHalfFinishNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DHalfFinishNoRqd", "DHalfFinishNo", "異常：需輸入半成品料號")
                DHalfFinishNo.Visible = True
            Case 2  '修改
                DHalfFinishNo.BackColor = Color.Yellow
                DHalfFinishNo.Visible = True
            Case Else   '隱藏
                DHalfFinishNo.Visible = False
        End Select
        If pPost = "New" Then DHalfFinishNo.Text = ""

        '加工種類 Material

        Select Case FindFieldInf("Material")
            Case 0  '顯示
                DMaterial.BackColor = Color.LightGray
                DMaterial.Visible = True
            Case 1  '修改+檢查
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "異常：需輸入半成品料號")
                DMaterial.Visible = True
            Case 2  '修改
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '隱藏
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")


        '加工種類 DMaterialDetailDetail 
        Select Case FindFieldInf("MaterialDetail")
            Case 0  '顯示
                DMaterialDetail.BackColor = Color.LightGray
                DMaterialDetail.Visible = True
            Case 1  '修改+檢查
                DMaterialDetail.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialDetailRqd", "DMaterialDetail", "異常：需輸入加工種類")
                DMaterialDetail.Visible = True
            Case 2  '修改
                DMaterialDetail.BackColor = Color.Yellow
                DMaterialDetail.Visible = True
            Case Else   '隱藏
                DMaterialDetail.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MaterialDetail", "ZZZZZZ")

        '加工種類 DMaterialDetailDetail 
        Select Case FindFieldInf("MaterialDetail_1")
            Case 0  '顯示
                DMaterialDetail_1.BackColor = Color.LightGray
                DMaterialDetail_1.Visible = True
                DMaterialDetail_1.ReadOnly = True
            Case 1  '修改+檢查
                DMaterialDetail_1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialDetail_1Rqd", "DMaterialDetail_1", "異常：加工種類備註")
                DMaterialDetail_1.Visible = True
            Case 2  '修改
                DMaterialDetail_1.BackColor = Color.Yellow
                DMaterialDetail_1.Visible = True
            Case Else   '隱藏
                DMaterialDetail_1.Visible = False
        End Select
        If pPost = "New" Then DMaterialDetail_1.Text = ""


        '草圖
        Select Case FindFieldInf("RefMapFile")

            Case 0  '顯示
                DRefMapFile.Visible = False
                DRefMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DRefMapFileRqd", "DRefMapFile", "異常：需輸入草圖")
                DRefMapFile.Visible = True
                DRefMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DRefMapFile.Visible = True
                DRefMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DRefMapFile.Visible = False
        End Select
        LRefMapFile.Visible = False


        '樣品
        Select Case FindFieldInf("Sample")
            Case 0  '顯示DSample
                DSample.BackColor = Color.LightGray
                DSample.Visible = True
            Case 1  '修改+檢查
                DSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleRqd", "DSample", "異常：需輸入樣品")
                DSample.Visible = True
            Case 2  '修改
                DSample.BackColor = Color.Yellow
                DSample.Visible = True
            Case Else   '隱藏
                DSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sample", "ZZZZZZ")

        '備註
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.ReadOnly = True
                DRemark.Visible = True
            Case 1  '修改+檢查
                DRemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需輸入備註")
                DRemark.Visible = True
            Case 2  '修改
                DRemark.BackColor = Color.Yellow
                DRemark.Visible = True
            Case Else   '隱藏
                DRemark.Visible = False
        End Select
        If pPost = "New" Then DRemark.Text = ""


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
        If pPost = "New" Then DMapNo.Text = ""


        '製圖者
        Select Case FindFieldInf("MakeMapU")
            Case 0  '顯示
                DMakeMapU.BackColor = Color.LightGray
                DMakeMapU.Visible = True
            Case 1  '修改+檢查
                DMakeMapU.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMakeMapURqd", "DMakeMapU", "異常：需輸入製圖者")
                DMakeMapU.Visible = True
            Case 2  '修改
                DMakeMapU.BackColor = Color.Yellow
                DMakeMapU.Visible = True
            Case Else   '隱藏
                DMakeMapU.Visible = False
        End Select


        '製圖者
        Select Case FindFieldInf("MakeMap")
            Case 0  '顯示
                DMakeMap.BackColor = Color.LightGray
                DMakeMap.Visible = True
            Case 1  '修改+檢查
                DMakeMap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMakeMapRqd", "DMakeMap", "異常：需輸入製圖者")
                DMakeMap.Visible = True
            Case 2  '修改
                DMakeMap.BackColor = Color.Yellow
                DMakeMap.Visible = True
            Case Else   '隱藏
                DMakeMap.Visible = False
        End Select


        If pPost = "New" Then SetFieldData("MakeMap", "ZZZZZZ")


        '難易度
        Select Case FindFieldInf("Level")
            Case 0  '顯示
                DLevel.BackColor = Color.LightGray
                DLevel.Visible = True

            Case 1  '修改+檢查
                DLevel.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLevelRqd", "DLevel", "異常：需輸入難易度")
                DLevel.Visible = True
            Case 2  '修改
                DLevel.BackColor = Color.Yellow
                DLevel.Visible = True
            Case Else   '隱藏
                DLevel.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")



        '圖檔
        Select Case FindFieldInf("MapFile")
            Case 0  '顯示
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DMapFile.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                If LMapFile.NavigateUrl <> "" Then
                    ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "異常：需輸入圖檔")
                End If

                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMapFile.Visible = False
        End Select

        LMapFile.Visible = False







        '原因類別
        Select Case FindFieldInf("Reason1")


            Case 0  '顯示
                DReason1.BackColor = Color.LightGray
                DReason1.Visible = True
            Case 1  '修改+檢查
                DReason1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReason1Rqd", "DReason1", "異常：需輸入原因類別")
                DReason1.Visible = True
            Case 2  '修改
                DReason1.BackColor = Color.Yellow
                DReason1.Visible = True
            Case Else   '隱藏
                DReason1.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("Reason1", "ZZZZZZ")


        '修改內容-1

        Dim sql As String
        Sql = "Select * From F_SBDCommissionSheet "
        Sql = Sql & " Where FormNo =  '" & wFormNo & "'"
        Sql = Sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(Sql)

        If DBAdapter1.Rows.Count > 0 Then
            flag = DBAdapter1.Rows(0).Item("Flag")
        End If




        Dim idx As Integer
        idx = FindFieldInf("1-Content")
        If flag > 0 And (wStep = 30) Then
            idx = 2
        End If

        Select Case idx
            Case 0  '顯示
                DContent1.BackColor = Color.LightGray
                DContent1.Visible = True
                DContent1.ReadOnly = True
                DContent2.BackColor = Color.LightGray
                DContent2.Visible = True
                DContent2.ReadOnly = True
                DContent3.BackColor = Color.LightGray
                DContent3.Visible = True
                DContent3.ReadOnly = True
                DContent4.BackColor = Color.LightGray
                DContent4.Visible = True
                DContent4.ReadOnly = True
                DContent5.BackColor = Color.LightGray
                DContent5.Visible = True
                DContent5.ReadOnly = True
                DContent6.BackColor = Color.LightGray
                DContent6.Visible = True
                DContent6.ReadOnly = True


            Case 1  '修改+檢查
                DContent1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent1", "異常：需輸入修改內容-1")
                DContent1.Visible = True
                DContent2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent2", "異常：需輸入修改內容-2")
                DContent2.Visible = True
                DContent3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent3", "異常：需輸入修改內容-3")
                DContent3.Visible = True
                DContent4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent4", "異常：需輸入修改內容-4")
                DContent4.Visible = True
                DContent5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent5", "異常：需輸入修改內容-5")
                DContent5.Visible = True
                DContent6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent6", "異常：需輸入修改內容-6")
                DContent6.Visible = True

            Case 2  '修改
                DContent1.BackColor = Color.Yellow
                DContent1.Visible = True
                DContent2.BackColor = Color.Yellow
                DContent2.Visible = True
                DContent3.BackColor = Color.Yellow
                DContent3.Visible = True
                DContent4.BackColor = Color.Yellow
                DContent4.Visible = True
                DContent5.BackColor = Color.Yellow
                DContent5.Visible = True
                DContent6.BackColor = Color.Yellow
                DContent6.Visible = True

            Case Else   '隱藏
                DContent1.Visible = False
                DContent2.Visible = False
                DContent3.Visible = False
                DContent4.Visible = False
                DContent5.Visible = False
                DContent6.Visible = False

        End Select
        If pPost = "New" Then DContent1.Text = ""
        If pPost = "New" Then DContent2.Text = ""
        If pPost = "New" Then DContent3.Text = ""
        If pPost = "New" Then DContent4.Text = ""
        If pPost = "New" Then DContent5.Text = ""
        If pPost = "New" Then DContent6.Text = ""





        idx = FindFieldInf("1-Contenfile")
        If flag > 0 And (wStep = 30) Then
            idx = 2
        End If
        '修改內容附檔
        Select Case idx
            Case 0  '顯示
                DContentFile1.Visible = False
                DContentFile1.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile1.Attributes.Add("readonly", "true")
                DContentFile2.Visible = False
                DContentFile2.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile2.Attributes.Add("readonly", "true")
                DContentFile3.Visible = False
                DContentFile3.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile3.Attributes.Add("readonly", "true")
                DContentFile4.Visible = False
                DContentFile4.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile4.Attributes.Add("readonly", "true")
                DContentFile5.Visible = False
                DContentFile5.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile5.Attributes.Add("readonly", "true")
                DContentFile6.Visible = False
                DContentFile6.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile6.Attributes.Add("readonly", "true")
                LContentFile1.Visible = True
                LContentFile1.BackColor = Color.LightGray
                LContentFile2.Visible = True
                LContentFile2.BackColor = Color.LightGray
                LContentFile3.Visible = True
                LContentFile3.BackColor = Color.LightGray
                LContentFile4.Visible = True
                LContentFile4.BackColor = Color.LightGray
                LContentFile5.Visible = True
                LContentFile5.BackColor = Color.LightGray
                LContentFile6.Visible = True
                LContentFile6.BackColor = Color.LightGray


            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DContentFile1Rqd", "DContentFile1", "異常：需輸入圖檔")
                DContentFile1.Visible = True
                DContentFile1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile2Rqd", "DContentFile2", "異常：需輸入圖檔")
                DContentFile2.Visible = True
                DContentFile2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile3Rqd", "DContentFile3", "異常：需輸入圖檔")
                DContentFile3.Visible = True
                DContentFile3.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile4Rqd", "DContentFile4", "異常：需輸入圖檔")
                DContentFile4.Visible = True
                DContentFile4.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile5Rqd", "DContentFile5", "異常：需輸入圖檔")
                DContentFile5.Visible = True
                DContentFile5.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile6Rqd", "DContentFile6", "異常：需輸入圖檔")
                DContentFile6.Visible = True
                DContentFile6.Style.Add("BACKGROUND-COLOR", "GreenYellow")

            Case 2  '修改
                DContentFile1.Visible = True
                DContentFile1.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile2.Visible = True
                DContentFile2.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile3.Visible = True
                DContentFile3.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile4.Visible = True
                DContentFile4.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile5.Visible = True
                DContentFile5.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile6.Visible = True
                DContentFile6.Style.Add("BACKGROUND-COLOR", "Yellow")

            Case Else   '隱藏
                DContentFile1.Visible = False
                DContentFile2.Visible = False
                DContentFile3.Visible = False
                DContentFile4.Visible = False
                DContentFile5.Visible = False
                DContentFile6.Visible = False
        End Select
        LContentFile1.Visible = False
        LContentFile2.Visible = False
        LContentFile3.Visible = False
        LContentFile4.Visible = False
        LContentFile5.Visible = False
        LContentFile6.Visible = False









        Select Case FindFieldInf("2-Makemap")
            Case 0  '顯示
                DMakeMap1.BackColor = Color.LightGray
                DMakeMap1.Visible = True
                DMakeMap2.BackColor = Color.LightGray
                DMakeMap2.Visible = True
                DMakeMap3.BackColor = Color.LightGray
                DMakeMap3.Visible = True
                DMakeMap4.BackColor = Color.LightGray
                DMakeMap4.Visible = True
                DMakeMap5.BackColor = Color.LightGray
                DMakeMap5.Visible = True
                DMakeMap6.BackColor = Color.LightGray
                DMakeMap6.Visible = True

                LMap1.Visible = True
                LMap2.Visible = True
                LMap3.Visible = True
                LMap4.Visible = True
                LMap5.Visible = True
                LMap6.Visible = True
            Case 1  '修改+檢查
                DMakeMap1.BackColor = Color.GreenYellow
                DMakeMap1.Visible = True
                DMakeMap2.BackColor = Color.GreenYellow
                DMakeMap2.Visible = True
                DMakeMap3.BackColor = Color.GreenYellow
                DMakeMap3.Visible = True
                DMakeMap4.BackColor = Color.GreenYellow
                DMakeMap4.Visible = True
                DMakeMap5.BackColor = Color.GreenYellow
                DMakeMap5.Visible = True
                DMakeMap6.BackColor = Color.GreenYellow
                DMakeMap6.Visible = True
                LMap1.Visible = True
                LMap2.Visible = True
                LMap3.Visible = True
                LMap4.Visible = True
                LMap5.Visible = True
                LMap6.Visible = True
            Case 2  '修改
                DMakeMap1.BackColor = Color.Yellow
                DMakeMap1.Visible = True
                DMakeMap2.BackColor = Color.Yellow
                DMakeMap2.Visible = True
                DMakeMap3.BackColor = Color.Yellow
                DMakeMap3.Visible = True
                DMakeMap4.BackColor = Color.Yellow
                DMakeMap4.Visible = True
                DMakeMap5.BackColor = Color.Yellow
                DMakeMap5.Visible = True
                DMakeMap6.BackColor = Color.Yellow
                DMakeMap6.Visible = True
                LMap1.Visible = True
                LMap2.Visible = True
                LMap3.Visible = True
                LMap4.Visible = True
                LMap5.Visible = True
                LMap6.Visible = True

            Case Else   '隱藏

                DMakeMap1.Visible = False
                DMakeMap2.Visible = False
                DMakeMap3.Visible = False
                DMakeMap4.Visible = False
                DMakeMap5.Visible = False
                DMakeMap6.Visible = False

                LMap1.Visible = False
                LMap2.Visible = False
                LMap3.Visible = False
                LMap4.Visible = False
                LMap5.Visible = False
                LMap6.Visible = False
        End Select











        '外注商
        Select Case FindFieldInf("Supplier")
            Case 0  '顯示
                DSupplier.BackColor = Color.LightGray
                DSupplier.Visible = True

            Case 1  '修改+檢查
                DSupplier.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSupplierRqd", "DSupplier", "異常：需輸入外注商")
                DSupplier.Visible = True
            Case 2  '修改
                DSupplier.BackColor = Color.Yellow
                DSupplier.Visible = True
            Case Else   '隱藏
                DSupplier.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Supplier", "ZZZZZZ")

        '半成品訂單NO.
        Select Case FindFieldInf("HalfFinishOrderNo")
            Case 0  '顯示
                DHalfFinishOrderNo.BackColor = Color.LightGray
                DHalfFinishOrderNo.Visible = True
                DHalfFinishOrderNo.ReadOnly = True
            Case 1  '修改+檢查
                DHalfFinishOrderNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DHalfFinishOrderNoRqd", "DHalfFinishOrderNo", "異常：需輸入半成品訂單NO.")
                DHalfFinishOrderNo.Visible = True
            Case 2  '修改
                DHalfFinishOrderNo.BackColor = Color.Yellow
                DHalfFinishOrderNo.Visible = True
            Case Else   '隱藏
                DHalfFinishOrderNo.Visible = False
        End Select
        If pPost = "New" Then DHalfFinishOrderNo.Text = ""

        '半成品預訂完成日

        Select Case FindFieldInf("HalfFinishDate")
            Case 0  '顯示
                DHalfFinishDate.Visible = True
                DHalfFinishDate.Style.Add("background-color", "lightgrey")
                DHalfFinishDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = True

            Case 1  '修改+檢查
                DHalfFinishDate.Visible = True
                DHalfFinishDate.Style.Add("background-color", "greenyellow")
                DHalfFinishDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DHalfFinishDateRqd", "DHalfFinishDate", "異常：需輸入半成品預訂完成日")
                BDate3.Disabled = False

            Case 2  '修改
                DHalfFinishDate.Visible = True
                DHalfFinishDate.Style.Add("background-color", "yellow")
                DHalfFinishDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = False

            Case Else   '隱藏
                DHalfFinishDate.Visible = False
                BDate3.Disabled = True

        End Select
        If pPost = "New" Then DHalfFinishDate.Value = ""

        '模具
        Select Case FindFieldInf("Mold")
            Case 0  '顯示
                DMold.BackColor = Color.LightGray
                DMold.Visible = True
                DMold.ReadOnly = True
            Case 1  '修改+檢查
                DMold.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldRqd", "DMold", "異常：需輸入模具")
                DMold.Visible = True
            Case 2  '修改
                DMold.BackColor = Color.Yellow
                DMold.Visible = True
            Case Else   '隱藏
                DMold.Visible = False
        End Select
        If pPost = "New" Then DMold.Text = ""



        '穴取
        Select Case FindFieldInf("MoldPoint")
            Case 0  '顯示
                DMoldPoint.BackColor = Color.LightGray
                DMoldPoint.Visible = True
                DMoldPoint.ReadOnly = True
            Case 1  '修改+檢查
                DMoldPoint.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldPointRqd", "DMoldPoint", "異常：需輸入穴取")
                DMoldPoint.Visible = True
            Case 2  '修改
                DMoldPoint.BackColor = Color.Yellow
                DMoldPoint.Visible = True
            Case Else   '隱藏
                DMoldPoint.Visible = False
        End Select
        If pPost = "New" Then DMoldPoint.Text = ""

        '表面顏色

        Select Case FindFieldInf("Surfcolor")
            Case 0  '顯示
                DSurfcolor.Visible = True
                DSurfcolor.Style.Add("background-color", "lightgrey")
                DSurfcolor.Attributes.Add("readonly", "true")
                BColor.Disabled = True
                If DSurfcolor.Value = "" Then
                    LSurfColor.Visible = False
                Else
                    LSurfColor.Visible = True
                End If


            Case 1  '修改+檢查
                DSurfcolor.Visible = True
                DSurfcolor.Style.Add("background-color", "greenyellow")
                DSurfcolor.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DSurfcolorRqd", "DSurfcolor", "異常：需輸入表面顏色")
                BColor.Disabled = False
                If DSurfcolor.Value = "" Then
                    LSurfColor.Visible = False
                Else
                    LSurfColor.Visible = True
                End If



            Case 2  '修改
                DSurfcolor.Visible = True
                DSurfcolor.Style.Add("background-color", "yellow")
                DSurfcolor.Attributes.Add("readonly", "true")
                BColor.Disabled = False
                If DSurfcolor.Value = "" Then
                    LSurfColor.Visible = False
                Else
                    LSurfColor.Visible = True
                End If


            Case Else   '隱藏
                DSurfcolor.Visible = False
                BColor.Disabled = True
                LSurfColor.Visible = False

        End Select
        If pPost = "New" Then DSurfcolor.Value = ""


        '表面顏色 1
        Select Case FindFieldInf("Surfcolor")
            Case 0  '顯示
                DSurfcolor1.BackColor = Color.LightGray
                DSurfcolor1.Visible = True
                DSurfcolor1.ReadOnly = True
            Case 1  '修改+檢查
                DSurfcolor1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSurfcolor1Rqd", "DSurfcolor1", "異常：表面顏色")
                DSurfcolor1.Visible = True
            Case 2  '修改
                DSurfcolor1.BackColor = Color.Yellow
                DSurfcolor1.Visible = True
            Case Else   '隱藏
                DSurfcolor1.Visible = False
        End Select
        If pPost = "New" Then DSurfcolor1.Text = ""



        '樣品需求量
        Select Case FindFieldInf("SampleQty")
            Case 0  '顯示
                DSampleQty.BackColor = Color.LightGray
                DSampleQty.Visible = True

            Case 1  '修改+檢查
                DSampleQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleQtyRqd", "DSampleQty", "異常：需輸入樣品需求量")
                DSampleQty.Visible = True
            Case 2  '修改
                DSampleQty.BackColor = Color.Yellow
                DSampleQty.Visible = True
            Case Else   '隱藏
                DSampleQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SampleQty", "ZZZZZZ")



        '分析依賴書
        Select Case FindFieldInf("QCReqFile")
            Case 0  '顯示
                DQCReqFile.Visible = False
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCReqFile.Attributes.Add("readonly", "true")
                LQCReqFile.Visible = True
                LQCReqFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQCReqFileRqd", "DQCReqFile", "異常：需輸入分析依賴書")
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQCReqFile.Visible = False
        End Select
        LQCReqFile.Visible = False

        'EA判定
        Select Case FindFieldInf("FQA")
            Case 0  '顯示
                DFQA.BackColor = Color.LightGray
                DFQA.Visible = True
                DFQA.ReadOnly = True

            Case 1  '修改+檢查
                DFQA.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFQARqd", "DFQA", "異常：需輸入EA判定")
                DFQA.Visible = True
            Case 2  '修改
                DFQA.BackColor = Color.Yellow
                DFQA.Visible = True
            Case Else   '隱藏
                DFQA.Visible = False
        End Select
        If pPost = "New" Then DFQA.Text = ""

        '品質備註
        Select Case FindFieldInf("QARemark")
            Case 0  '顯示
                DQARemark.BackColor = Color.LightGray
                DQARemark.Visible = True
                DQARemark.ReadOnly = True
            Case 1  '修改+檢查
                DQARemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQARemarkRqd", "DQARemark", "異常：需輸入品質備註")
                DQARemark.Visible = True
            Case 2  '修改
                DQARemark.BackColor = Color.Yellow
                DQARemark.Visible = True
            Case Else   '隱藏
                DQARemark.Visible = False
        End Select
        If pPost = "New" Then DQARemark.Text = ""

        '報價單
        Select Case FindFieldInf("ForcastFile")
            Case 0  '顯示
                DForcastFile.Visible = False
                DForcastFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DForcastFile.Attributes.Add("readonly", "true")
                LForcastFile.Visible = True
                LForcastFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DForcastFileRqd", "DForcastFile", "異常：需輸入報價單")
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DForcastFile.Visible = False
        End Select
        LForcastFile.Visible = False

        '品測報告書
        Select Case FindFieldInf("QAFile")
            Case 0  '顯示
                DQAFile.Visible = False
                DQAFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQAFile.Attributes.Add("readonly", "true")
                LQAFile.Visible = True
                LQAFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQAFileRqd", "DQAFile", "異常：需輸入品測報告書")
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DQAFile.Visible = False
        End Select
        LQAFile.Visible = False


        '確認書
        Select Case FindFieldInf("AuthorizeFile")
            Case 0  '顯示
                DAuthorizeFile.Visible = False
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DAuthorizeFile.Attributes.Add("readonly", "true")
                LAuthorizeFile.Visible = True
                LAuthorizeFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAuthorizeFileRqd", "DAuthorizeFile", "異常：需輸入確認書")
                DAuthorizeFile.Visible = True
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DAuthorizeFile.Visible = True
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DAuthorizeFile.Visible = False
        End Select
        LAuthorizeFile.Visible = False


        '最終樣品圖
        Select Case FindFieldInf("SampleFile")
            Case 0  '顯示
                DSampleFile.Visible = False
                DSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DSampleFile.Attributes.Add("readonly", "true")
                LSampleFile.Visible = True
                LSampleFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "異常：需輸入最終樣品圖")
                DSampleFile.Visible = True
                DSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DSampleFile.Visible = True
                DSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DSampleFile.Visible = False
        End Select
        LSampleFile.Visible = False


        '切結書
        Select Case FindFieldInf("ContactFile")
            Case 0  '顯示
                DContactFile.Visible = False
                DContactFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContactFile.Attributes.Add("readonly", "true")
                LContactFile.Visible = True
                LContactFile.BackColor = Color.LightGray

            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DContactFileRqd", "DContactFile", "異常：需輸入切結書")
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DContactFile.Visible = False
        End Select
        LContactFile.Visible = False


        '其它文件
        Select Case FindFieldInf("RefFile")
            Case 0  '顯示
                DRefFile.Visible = False
                DRefFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DRefFile.Attributes.Add("readonly", "true")
                LRefFile.Visible = True
                LRefFile.BackColor = Color.LightGray
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DRefFileRqd", "DRefFile", "異常：需輸入其它文件")
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DRefFile.Visible = False
        End Select
        LRefFile.Visible = False



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
        DAppper.Text = DBUser.Rows(0).Item("Username")
        DDivision.Text = DBUser.Rows(0).Item("Divname")

        '光造型
        If pFieldName = "Light" Then
            DLight.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLight.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Light' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DLight.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLight.Items.Add(ListItem1)
                Next
            End If
        End If

        'buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBuyer.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Buyer' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DBuyer.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBuyer.Items.Add(ListItem1)
                Next
            End If
        End If



        '加工種類
        If pFieldName = "Material" Then
            DMaterial.Items.Clear()
            DMaterialDetail.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMaterial.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'MATERIAL' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DMaterial.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMaterial.Items.Add(ListItem1)
                Next
            End If
        End If



        '加工種類
        If pFieldName = "MaterialDetail" Then
            DMaterialDetail.Items.Clear()
            DMaterialDetail.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMaterialDetail.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3002' "
                SQL = SQL & "   And dkey = '" + DMaterial.SelectedValue + "'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DMaterialDetail.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMaterialDetail.Items.Add(ListItem1)
                Next
            End If
        End If






        '附樣品
        If pFieldName = "Sample" Then
            DSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Sample'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSample.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSample.Items.Add(ListItem1)
                Next
            End If
        End If

        '樣品需求量
        If pFieldName = "SampleQty" Then
            DSampleQty.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSampleQty.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'SampleQty' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSampleQty.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSampleQty.Items.Add(ListItem1)
                Next
            End If
        End If


        '製圖者
        If pFieldName = "MakeMap" Then

            DMakeMap.Items.Clear()


            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMakeMap.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'MakeMap'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DMakeMap.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMakeMap.Items.Add(ListItem1)


                Next
            End If
        End If






        'LEVEL
        If pFieldName = "Level" Then
            DLevel.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLevel.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Level'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DLevel.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLevel.Items.Add(ListItem1)
                Next
            End If
        End If

        '原因類別
        If pFieldName = "Reason1" Then
            DReason1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReason1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Reason'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DReason1.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReason1.Items.Add(ListItem1)
                Next
            End If
        End If

        '外注商
        If pFieldName = "Supplier" Then
            DSupplier.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSupplier.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Supplier'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSupplier.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSupplier.Items.Add(ListItem1)
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

                    Dim URL As String = uCommon.GetAppSetting("RedirectURL") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
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
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If wStep = 70 Then
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
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")
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
                ErrCode = oCommon.CommissionNo("003001", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
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
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End
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

                    Dim wAllocateID As String = ""
                    PAction1 = pAction '為了60工程
                    Select Case wStep
                        Case 30
                            If pAction = 0 Then  '2013/01/10   jessica
                                wAllocateID = oCommon.GetUserID(DMakeMap.Text)
                            End If
                        Case 50
                            If pAction = 1 Then  '2013/01/10   jessica
                                If DMakeMap.SelectedValue = "" Then
                                    If flag = 0 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMapU.Text)
                                    ElseIf flag = 1 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap1.Text)
                                    ElseIf flag = 2 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap2.Text)
                                    ElseIf flag = 3 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap3.Text)
                                    ElseIf flag = 4 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap4.Text)
                                    ElseIf flag = 5 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap5.Text)
                                    ElseIf flag = 6 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap5.Text)
                                    End If
                                Else
                                    wAllocateID = oCommon.GetUserID(DMakeMap.SelectedValue)
                                End If
                               
                            End If
                        Case Else

                    End Select

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
                            Select Case wStep
                                Case 60
                                    If PAction1 = 1 Then
                                        '更新flag 重新開始
                                        SQL = " update f_sbdcommissionsheet "
                                        SQL = SQL + " Set flag = 0"
                                        SQL = SQL + " ,mapno=''"
                                        SQL = SQL + " ,makemap =''"
                                        SQL = SQL + " ,mapfile = ''"
                                        SQL = SQL + " ,map1 =''"
                                        SQL = SQL + " ,makemap1 =''"
                                        SQL = SQL + " ,map2 = ''"
                                        SQL = SQL + " ,makemap2 = ''"
                                        SQL = SQL + " ,map3 = ''"
                                        SQL = SQL + " ,makemap3 = ''"
                                        SQL = SQL + " ,map4 = ''"
                                        SQL = SQL + " ,makemap4 = ''"
                                        SQL = SQL + " ,map5 = ''"
                                        SQL = SQL + " ,makemap5 = ''"
                                        SQL = SQL + " ,ModifyUser = '" & Request.QueryString("pUserID") & "',"
                                        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                                        SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                                        SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                                        uDataBase.ExecuteNonQuery(SQL)

                                    End If
                                Case Else

                            End Select
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
                                                                 "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
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
            If ErrCode = 45 Then Message = "異常：需輸入外注商!"
            If ErrCode = 46 Then Message = "異常：需輸入半成品訂單No!"
            If ErrCode = 47 Then Message = "異常：需輸入半成品預訂完成日!"
            If ErrCode = 48 Then Message = "異常：需輸入模具!"
            If ErrCode = 49 Then Message = "異常：需輸入模型!"
            If ErrCode = 50 Then Message = "異常：需輸入穴取!"

            If ErrCode = 51 Then Message = "異常：需輸入表面顏色!"
            If ErrCode = 52 Then Message = "異常：需輸入樣品數量!"
            If ErrCode = 53 Then Message = "異常：需輸入分析依賴書!"
            If ErrCode = 54 Then Message = "異常：需輸入EA判定!"
            If ErrCode = 55 Then Message = "異常：需輸入品質備註!"

            If ErrCode = 56 Then Message = "異常：需輸入報價單!"
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
                   System.Configuration.ConfigurationManager.AppSettings("SBDCommissionFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String



        SQl = "Insert into F_SBDCommissionSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Appdate, Buyer, Vendor, Division, "                  '1~5
        SQl = SQl + "APPPER, Background,Code,Mapdate, Sampledate, "               '6~10
        SQl = SQl + "Light,halffinishNo, material, MaterialDetail,MaterialDetail_1, "                 '11~15
        SQl = SQl + "RefMapFile, Sample, Remark, Mapno, MakeMap,"   '16~20
        SQl = SQl + "level,MapFile,Reason1, "                                                       '21~23
        SQl = SQl + "Content1,Content2,Content3,Content4,Content5,Content6, "            '24~29
        SQl = SQl + "Contentfile1,Contentfile2,Contentfile3,Contentfile4,Contentfile5,Contentfile6, "            '30~35
        SQl = SQl + "Map1,Map2,Map3,Map4,Map5,Map6, "            '36~41
        SQl = SQl + "Makemap1,Makemap2,Makemap3,Makemap4,Makemap5,Makemap6, "            '42~47
        SQl = SQl + "Supplier,HalfFinishOrderNo,HalfFinishDate,Mold,"            '48~52
        SQl = SQl + "MoldPoint,Surfcolor,Surfcolor1,SurfFormNo,SurfFormSno,SampleQty,QCReqFile,FQA, "            '53~57
        SQl = SQl + "QARemark,ForcastFile,QAFile,AuthorizeFile,SampleFile, "            '58~62
        SQl = SQl + "ContactFile,RefFile, "            '63~64
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "


        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '003001', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號

        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        SQl = SQl + " '" + DAppDate.Text + "', "                '日期
        SQl = SQl + " '" + DBuyer.SelectedValue + "', "   'BUYER
        SQl = SQl + " '" + DVendor.Text + "', "     'VENDOR
        SQl = SQl + " '" + DDivision.Text + "', "  '部門

        SQl = SQl + " '" + DAppper.Text + "', "    '擔當
        SQl = SQl + " '" + DBackground.Text + "', "      '開發背景
        SQl = SQl + " N'" + YKK.ReplaceString(DCode.Text) + "', "   '料號
        SQl = SQl + " '" + DMapDate.Value + "', "     '圖面希望交期
        SQl = SQl + " '" + DSampleDate.Value + "', "  '樣品希望交期


        SQl = SQl + " '" + DLight.SelectedValue + "', "    '光造型
        SQl = SQl + " '" + DHalfFinishNo.Text + "',"   '半成品料
        SQl = SQl + " '" + DMaterial.SelectedValue + "', "   '加工種類-1
        SQl = SQl + " '" + DMaterialDetail.SelectedValue + "', "   '加工種類-2
        SQl = SQl + " '" + DMaterialDetail_1.Text + "', "   '加工種類-3
        'FileName = ""
        ' If DRefMapFile.PostedFile.FileName <> "" Then

        'FileName = UploadDateTime & "-" & Right(DRefMapFile.PostedFile.FileName, InStr(StrReverse(DRefMapFile.PostedFile.FileName), "\") - 1)
        'DRefMapFile.PostedFile.SaveAs(Path + FileName)
        'Else
        'FileName = ""
        'End If


        FileName = ""
        If DRefMapFile.Visible Then
            If DRefMapFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "RefMapFile" & "-" & UploadDateTime & "-" & Right(DRefMapFile.PostedFile.FileName, InStr(StrReverse(DRefMapFile.PostedFile.FileName), "\") - 1)

                DRefMapFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If




        SQl = SQl + " '" + FileName + "'," '草圖
        SQl = SQl + " '" + DSample.Text + "', "   '樣品
        SQl = SQl + " '" + DRemark.Text + "', "    '備註 
        SQl = SQl + " '" + DMapNo.Text + "', "    '圖號
        SQl = SQl + " '" + DMakeMap.SelectedValue + "', "    '製圖者    16~20

        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                DMapFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + DLevel.Text + "', "    '難易度
        SQl = SQl + " '" + FileName + "'," '圖檔
        SQl = SQl + " '" + DReason1.SelectedValue + "', "    '原因類別 

        SQl = SQl + " '" + DContent1.Text + "', "    '修改內容1
        SQl = SQl + " '" + DContent2.Text + "', "    '修改內容2
        SQl = SQl + " '" + DContent3.Text + "', "    '修改內容3
        SQl = SQl + " '" + DContent4.Text + "', "    '修改內容4
        SQl = SQl + " '" + DContent5.Text + "', "    '修改內容5
        SQl = SQl + " '" + DContent6.Text + "', "    '修改內容6

        FileName = ""

        If DContentFile1.Visible Then
            If DContentFile1.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile1" & "-" & UploadDateTime & "-" & Right(DContentFile1.PostedFile.FileName, InStr(StrReverse(DContentFile1.PostedFile.FileName), "\") - 1)
                DContentFile1.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '修改內容附檔1

        FileName = ""
        If DContentFile2.Visible Then
            If DContentFile2.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile2" & "-" & UploadDateTime & "-" & Right(DContentFile2.PostedFile.FileName, InStr(StrReverse(DContentFile2.PostedFile.FileName), "\") - 1)
                DContentFile2.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '修改內容附檔2

        FileName = ""
        If DContentFile3.Visible Then
            If DContentFile3.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile3" & "-" & UploadDateTime & "-" & Right(DContentFile3.PostedFile.FileName, InStr(StrReverse(DContentFile3.PostedFile.FileName), "\") - 1)
                DContentFile3.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '修改內容附檔3

        FileName = ""

        If DContentFile4.Visible Then
            If DContentFile4.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile4" & "-" & UploadDateTime & "-" & Right(DContentFile4.PostedFile.FileName, InStr(StrReverse(DContentFile4.PostedFile.FileName), "\") - 1)
                DContentFile4.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '修改內容附檔4

        FileName = ""

        If DContentFile5.Visible Then
            If DContentFile5.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile5" & "-" & UploadDateTime & "-" & Right(DContentFile5.PostedFile.FileName, InStr(StrReverse(DContentFile5.PostedFile.FileName), "\") - 1)
                DContentFile5.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '修改內容附檔5

        FileName = ""

        If DContentFile6.Visible Then
            If DContentFile6.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile6" & "-" & UploadDateTime & "-" & Right(DContentFile6.PostedFile.FileName, InStr(StrReverse(DContentFile6.PostedFile.FileName), "\") - 1)
                DContentFile6.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '修改內容附檔6



        SQl = SQl + " ''," '圖檔1
        SQl = SQl + " ''," '圖檔2
        SQl = SQl + " ''," '圖檔3
        SQl = SQl + " ''," '圖檔4
        SQl = SQl + " ''," '圖檔5
        SQl = SQl + " ''," '圖檔6

        SQl = SQl + " ''," '製圖1
        SQl = SQl + " ''," '製圖2
        SQl = SQl + " ''," '製圖3
        SQl = SQl + " ''," '製圖4
        SQl = SQl + " ''," '製圖5
        SQl = SQl + " ''," '製圖6



        SQl = SQl + " '" + DSupplier.SelectedValue + "', "    '外注商
        SQl = SQl + " '" + DHalfFinishOrderNo.Text + "', "    '半成品訂單NO
        SQl = SQl + " '" + DHalfFinishDate.Value + "', "      '半成品預訂完成日期
        SQl = SQl + " '" + DMold.Text + "', "                 '模具
        '
        SQl = SQl + " '" + DMoldPoint.Text + "', "            '穴取
        SQl = SQl + " '" + DSurfcolor.Value + "', "            '表面顏色
        SQl = SQl + " '" + DSurfcolor1.Text + "', "            '表面顏色
        SQl = SQl + " '', "            '表面顏色
        SQl = SQl + " '', "            '表面顏色

        SQl = SQl + " '" + DSampleQty.SelectedValue + "', "            '樣品數量

        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                DQCReqFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '分析依賴書

        SQl = SQl + " '" + DFQA.Text + "', "                  'EA判定
        SQl = SQl + " '" + DQARemark.Text + "', "             '品質備註

        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If


        SQl = SQl + " '" + FileName + "',"                    '報價單

        FileName = ""

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                DQAFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '品測報告書

        FileName = ""
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "AuthorizeFile" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '確認書

        FileName = ""

        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                DSampleFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '最終樣品圖


        FileName = ""
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContactFile" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                DContactFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '切結書

        FileName = ""
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "RefFile" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                DRefFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '其它文件



        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
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
               System.Configuration.ConfigurationManager.AppSettings("SBDCommissionFilePath")
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String



        ' SQl = "Select Divname,Username From M_Users ")
        ' SQl = SQl & " Where UserID = '" & Request.QueryString("pUserID") & "'"
        ' SQl = SQl & "   And Active = '1' "
        ' Dim DBUser As DataTable = uDataBase.GetDataTable(SQl)
        'DMakeMap.SelectedValue = DBUser.Rows(0).Item("Username")



        '算已經修改幾次
        Dim sql2 As String = ""
        sql2 = " select flag from F_SBDCommissionSheet "
        sql2 = sql2 + " where formno =  '" & wFormNo & "'"
        sql2 = sql2 + " and formsno = '" & CStr(wFormSno) & "'"
        Dim DBData As DataTable = uDataBase.GetDataTable(sql2)
        Dim sflag As String = ""
        flag = DBData.Rows(0).Item("flag")





        SQl = "Update F_SBDCommissionSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"

        SQl = SQl + " AppDate = '" & DAppDate.Text & "',"
        SQl = SQl + " Buyer = '" & DBuyer.SelectedValue & "',"
        SQl = SQl + " Vendor = '" & DVendor.Text & "',"
        SQl = SQl + " Division = '" & DDivision.Text & "',"
        SQl = SQl + " APPPER = '" & DAppper.Text & "',"

        SQl = SQl + " Background = '" & DBackground.Text & "',"
        SQl = SQl + " Code = '" & DCode.Text & "',"
        SQl = SQl + " MapDate = '" & DMapDate.Value & "',"
        SQl = SQl + " SampleDate = '" & DSampleDate.Value & "',"
        SQl = SQl + " Light = '" & DLight.SelectedValue & "',"

        SQl = SQl + " HalfFinishNo = '" & DHalfFinishNo.Text & "',"
        SQl = SQl + " Material= '" & DMaterial.SelectedValue & "',"
        SQl = SQl + " MaterialDetail= '" & DMaterialDetail.SelectedValue & "',"
        SQl = SQl + " MaterialDetail_1= '" & DMaterialDetail_1.Text & "',"

        If wStep = 50 And PAction1 = 1 Then
            SQl = SQl + " flag = flag +1" & ","
        End If

        FileName = ""
        If DRefMapFile.Visible Then
            If DRefMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "RefMapFile" & "-" & UploadDateTime & "-" & Right(DRefMapFile.PostedFile.FileName, InStr(StrReverse(DRefMapFile.PostedFile.FileName), "\") - 1)
                DRefMapFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " RefMapFile= N'" + YKK.ReplaceString(FileName) + "',"           '草圖
            Else
                FileName = ""
            End If
        End If



        SQl = SQl + " Sample= '" & DSample.SelectedValue & "',"                      '樣品
        SQl = SQl + " Remark= '" & DRemark.Text & "',"                               '備註
        SQl = SQl + " MapNo= '" & DMapNo.Text & "',"                                 '圖號
        If wStep = 30 Then
            If DMakeMap.SelectedValue <> "" Then
                SQl = SQl + " MakeMap= '" & DMakeMap.SelectedValue & "',"
            End If
            '製圖者
        End If


        SQl = SQl + " Level= '" & DLevel.Text & "',"                                 '難易度



        If flag = 0 Then
            FileName = ""
            If DMapFile.Visible Then
                If DMapFile.PostedFile.FileName <> "" Then

                    FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " MapFile= N'" + YKK.ReplaceString(FileName) + "',"           '圖檔
                Else
                    FileName = ""
                End If
            End If

        End If




        SQl = SQl + " Reason1= '" & DReason1.SelectedValue & "',"                   '原因類別

        SQl = SQl + " Content1= '" + DContent1.Text + "', "    '修改內容1
        SQl = SQl + " Content2= '" + DContent2.Text + "', "    '修改內容2
        SQl = SQl + " Content3= '" + DContent3.Text + "', "    '修改內容3
        SQl = SQl + " Content4= '" + DContent4.Text + "', "    '修改內容4
        SQl = SQl + " Content5= '" + DContent5.Text + "', "    '修改內容5
        SQl = SQl + " Content6= '" + DContent6.Text + "', "    '修改內容6

        FileName = ""
        If DContentFile1.Visible Then
            If DContentFile1.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile1" & "-" & UploadDateTime & "-" & Right(DContentFile1.PostedFile.FileName, InStr(StrReverse(DContentFile1.PostedFile.FileName), "\") - 1)
                DContentFile1.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContentFile1=  N'" + YKK.ReplaceString(FileName) + "'," '修改內容附檔1
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DContentFile2.Visible Then
            If DContentFile2.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile2" & "-" & UploadDateTime & "-" & Right(DContentFile2.PostedFile.FileName, InStr(StrReverse(DContentFile2.PostedFile.FileName), "\") - 1)
                DContentFile2.PostedFile.SaveAs(Path + FileName)

                SQl = SQl + " ContentFile2=  N'" + YKK.ReplaceString(FileName) + "'," '修改內容附檔2
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DContentFile3.Visible Then
            If DContentFile3.PostedFile.FileName <> "" Then
                FileName = CStr(wFormSno) & "-" & "ContentFile3" & "-" & UploadDateTime & "-" & Right(DContentFile3.PostedFile.FileName, InStr(StrReverse(DContentFile3.PostedFile.FileName), "\") - 1)
                DContentFile3.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContentFile3=  N'" + YKK.ReplaceString(FileName) + "'," '修改內容附檔3
                FileName = ""
            End If
        End If



        FileName = ""
        If DContentFile4.Visible Then
            If DContentFile4.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile4" & "-" & UploadDateTime & "-" & Right(DContentFile4.PostedFile.FileName, InStr(StrReverse(DContentFile4.PostedFile.FileName), "\") - 1)
                DContentFile4.PostedFile.SaveAs(Path + FileName)

                SQl = SQl + " ContentFile4=  N'" + YKK.ReplaceString(FileName) + "'," '修改內容附檔4
            Else
                FileName = ""
            End If
        End If


        FileName = ""

        If DContentFile5.Visible Then
            If DContentFile5.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile5" & "-" & UploadDateTime & "-" & Right(DContentFile5.PostedFile.FileName, InStr(StrReverse(DContentFile5.PostedFile.FileName), "\") - 1)
                DContentFile5.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContentFile5=  N'" + YKK.ReplaceString(FileName) + "'," '修改內容附檔5
            Else
                FileName = ""
            End If
        End If



        FileName = ""

        If DContentFile6.Visible Then
            If DContentFile6.PostedFile.FileName <> "" Then
                FileName = CStr(wFormSno) & "-" & "ContentFile6" & "-" & UploadDateTime & "-" & Right(DContentFile6.PostedFile.FileName, InStr(StrReverse(DContentFile6.PostedFile.FileName), "\") - 1)
                DContentFile6.PostedFile.SaveAs(Path + FileName)

                SQl = SQl + " ContentFile6= N'" + YKK.ReplaceString(FileName) + "'," '修改內容附檔6
            Else
                FileName = ""
            End If
        End If


        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                If flag = 1 Then
                    FileName = CStr(wFormSno) & "-" & "Map1" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map1= N'" + YKK.ReplaceString(FileName) + "'," '圖檔2

                ElseIf flag = 2 Then
                    FileName = CStr(wFormSno) & "-" & "Map2" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map2= N'" + YKK.ReplaceString(FileName) + "'," '圖檔2

                ElseIf flag = 3 Then
                    FileName = CStr(wFormSno) & "-" & "Map3" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map3= N'" + YKK.ReplaceString(FileName) + "'," '圖檔3

                ElseIf flag = 4 Then
                    FileName = CStr(wFormSno) & "-" & "Map4" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map4= N'" + YKK.ReplaceString(FileName) + "'," '圖檔4

                ElseIf flag = 5 Then
                    FileName = CStr(wFormSno) & "-" & "Map5" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map5= N'" + YKK.ReplaceString(FileName) + "'," '圖檔5

                ElseIf flag = 6 Then
                    FileName = CStr(wFormSno) & "-" & "Map6" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map6= N'" + YKK.ReplaceString(FileName) + "'," '圖檔6


                End If

            End If
        End If


        If wStep = 50 And PAction1 = 1 Then
            If DMakeMap.SelectedValue = "" Then
                If flag = 0 Then
                    SQl = SQl + " Makemap1= '" + DMakeMapU.Text + "', "    '製圖者1 
                ElseIf flag = 1 Then
                    SQl = SQl + " Makemap2= '" + DMakeMap1.Text + "', "    '製圖者1 
                ElseIf flag = 2 Then
                    SQl = SQl + " Makemap3= '" + DMakeMap2.Text + "', "    '製圖者1 
                ElseIf flag = 3 Then
                    SQl = SQl + " Makemap4= '" + DMakeMap3.Text + "', "    '製圖者1 
                ElseIf flag = 4 Then
                    SQl = SQl + " Makemap5= '" + DMakeMap4.Text + "', "    '製圖者1 
                ElseIf flag = 5 Then
                    SQl = SQl + " Makemap6= '" + DMakeMap5.Text + "', "    '製圖者1 
                End If
            Else
                If flag = 0 Then
                    SQl = SQl + " Makemap1= '" + DMakeMap.SelectedValue + "', "    '製圖者1 
                ElseIf flag = 1 Then
                    SQl = SQl + " Makemap2= '" + DMakeMap.SelectedValue + "', "    '製圖者1 
                ElseIf flag = 2 Then
                    SQl = SQl + " Makemap3= '" + DMakeMap.SelectedValue + "', "    '製圖者1 
                ElseIf flag = 3 Then
                    SQl = SQl + " Makemap4= '" + DMakeMap.SelectedValue + "', "    '製圖者1 
                ElseIf flag = 4 Then
                    SQl = SQl + " Makemap5= '" + DMakeMap.SelectedValue + "', "    '製圖者1 
                ElseIf flag = 5 Then
                    SQl = SQl + " Makemap6= '" + DMakeMap.SelectedValue + "', "    '製圖者1 
                End If
            End If


        End If


        SQl = SQl + " Supplier= '" + DSupplier.SelectedValue + "', "    '外注商
        SQl = SQl + " HalfFinishOrderNo= '" + DHalfFinishOrderNo.Text + "', "    '半成品訂單NO
        SQl = SQl + " HalfFinishdate= '" + DHalfFinishDate.Value + "', "      '半成品預訂完成日期
        SQl = SQl + " Mold= '" + DMold.Text + "', "                 '模具
        '
        SQl = SQl + " MoldPoint= '" + DMoldPoint.Text + "', "            '穴取



        SQl = SQl + " Surfcolor= '" + DSurfcolor.Value + "', "            '表面顏色


        If DSurfFormNo.Text <> "" Then
            SQl = SQl + " SurfFormNo= '" + DSurfFormNo.Text + "', "            '表面 formno
            SQl = SQl + " SurfFormSno= '" + DSurfFormSno.Text + "', "            '表面顏色 formsno
        End If

        SQl = SQl + " Surfcolor1= '" + DSurfcolor1.Text + "', "            '表面顏色1

        SQl = SQl + " SampleQty= '" + DSampleQty.SelectedValue + "', "            '樣品數量



        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                DQCReqFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QCReqFile=   N'" + YKK.ReplaceString(FileName) + "'," '分析依賴書
            Else
                FileName = ""
            End If
        End If



        SQl = SQl + " FQA= '" + DFQA.Text + "', "                  'EA判定
        SQl = SQl + " QARemark= '" + DQARemark.Text + "', "             '品質備註

        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ForcastFile= N'" + YKK.ReplaceString(FileName) + "',"                 '報價單
            Else
                FileName = ""
            End If
        End If



        FileName = ""

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                DQAFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QAFile= N'" + YKK.ReplaceString(FileName) + "',"                 '品測報告書

            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "AuthorizeFile" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " AuthorizeFile=  N'" + YKK.ReplaceString(FileName) + "',"                   '確認書
            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                DSampleFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " SampleFile=  N'" + YKK.ReplaceString(FileName) + "',"                   '最終樣品圖
            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContactFile" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                DContactFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContactFile= N'" + YKK.ReplaceString(FileName) + "',"                     '切結書
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "RefFile" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                DRefFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " RefFile= N'" + YKK.ReplaceString(FileName) + "',"                     '其它文件
            Else
                FileName = ""
            End If
        End If






        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(SQl)
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
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
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
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function OK() As Boolean

        Dim isOK As Boolean = True
        Dim Errcode As Integer = 0



        Dim Message As String = ""


        '外注商
        If Errcode = 0 Then
            If DSupplier.Text = "" Then Errcode = 45
        End If


        '預訂完成日
        If Errcode = 0 Then
            If DHalfFinishDate.Value = "" Then Errcode = 47
        End If

        '模具
        If Errcode = 0 Then

            If DMold.Text = "" Then Errcode = 48

        End If



        '穴取
        If Errcode = 0 Then
            If DMoldPoint.Text = "" Then Errcode = 50

        End If

        '表面顏色
        If Errcode = 0 Then
            If DSurfcolor.Value = "" And DSurfcolor1.Text = "" Then Errcode = 51

        End If

        '樣品數量
        If Errcode = 0 Then

            If DSampleQty.SelectedValue = "" Then Errcode = 52

        End If



        '報價單 
        If Errcode = 0 Then

            If LForcastFile.Text = "" Then Errcode = 56

        End If






        If Errcode > 0 Then
            isOK = False
            If Errcode = 45 Then Message = "異常：需輸入外注商!"
            If Errcode = 46 Then Message = "異常：需輸入半成品訂單No!"
            If Errcode = 47 Then Message = "異常：需輸入半成品預訂完成日!"
            If Errcode = 48 Then Message = "異常：需輸入模具!"
            If Errcode = 49 Then Message = "異常：需輸入模型!"
            If Errcode = 50 Then Message = "異常：需輸入穴取!"

            If Errcode = 51 Then Message = "異常：需輸入表面顏色!"
            If Errcode = 52 Then Message = "異常：需輸入樣品數量!"
            If Errcode = 53 Then Message = "異常：需輸入分析依賴書!"
            If Errcode = 54 Then Message = "異常：需輸入EA判定!"
            If Errcode = 55 Then Message = "異常：需輸入品質備註!"

            If Errcode = 56 Then Message = "異常：需輸入報價單!"

        End If

        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If


        Return isOK


    End Function

    Protected Sub DMaterial_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMaterial.SelectedIndexChanged
        '加工種類2
        Dim Sql As String
        Dim i As Integer

        DMaterialDetail.Items.Clear()

        Sql = "Select Data From M_referp "
        Sql = Sql & " Where cat  = '3002' "
        Sql = Sql & "   And dkey = '" + DMaterial.SelectedValue + "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(Sql)
        DMaterialDetail.Items.Add("")

        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
            ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
            DMaterialDetail.Items.Add(ListItem1)
        Next


    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()
        TopPosition()
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

    Function InputCheck() As Integer

    End Function

    Protected Sub DMakeMap_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMakeMap.SelectedIndexChanged

    End Sub
End Class

