Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls


Partial Class ComplaintOutSheet_01
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
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String
    Dim chktempFolderPath As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"
        ' BQCIDATE1.Attributes("onclick") = "calendarPicker('Form1.DAppDate');"
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not Me.IsPostBack Then   '不是PostBack
            NewAttachFilePath()

            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                NewAttachFilePath2()
                NewAttachFilePath3()
                NewAttachFilePath4()
                NewAttachFilePath5()
                NewAttachFilePath6()

                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件


        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

            '上傳資料檢查及顯示訊息

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        If wFormSno > 0 And wStep > 3 Then      '判斷是否[簽核]
            Top = 2050
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 2050
                    End If
                End If
            End If
        Else
            Top = 2050

        End If
    End Sub


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
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()


        Top = 2050
        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 85 & "px"
            DDecideDesc.Style("top") = Top - 80 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If
        If wStep > 0 Then
            Top = 2050
        End If


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"


        Top = Top + 50
        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
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
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        '
        Response.Cookies("PGM").Value = "FinalcheckSheet_01.aspx"
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
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim sql As String = ""
        sql = "Select * From M_Flow "
        sql &= " Where Active = 1 "
        sql &= "   And FormNo =  '" & wFormNo & "'"
        sql &= "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)
        If dtFlow.Rows.Count > 0 Then
            '電子簽章未使用
            If dtFlow.Rows(0)("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If dtFlow.Rows(0)("Attach") = 1 Then
            Else
            End If
            '儲存按鈕
            If dtFlow.Rows(0)("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Text = dtFlow.Rows(0)("SaveDesc")
                'BSAVE.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSAVE.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                'BNG1.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG1.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                'BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                'BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If dtFlow.Rows(0)("Delay") = 1 Then
                wDelay = 1
            End If
        End If
        '
        If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
            sql = "Select * From T_WaitHandle "
            sql = sql & " Where Active = 1 "
            sql = sql & "   And FormNo =  '" & wFormNo & "'"
            sql = sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
            sql = sql & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
            If dtWaitHandle.Rows.Count > 0 Then
                '說明設定
                Top = 2050
                DDescSheet.Visible = True
                DDecideDesc.Visible = True
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                '遲納管理
                If wDelay = 1 Then
                    If dtWaitHandle.Rows(0)("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                        DReasonCode.Visible = True     '延遲理由代碼
                        DReasonCode.BackColor = Color.GreenYellow
                        DReason.Visible = True         '延遲理由
                        DReason.BackColor = Color.GreenYellow
                        DReasonDesc.Visible = True     '延遲其他說明
                        ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                        ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                        Top = Top + 110
                    Else
                        DDelay.Visible = False  '延遲Sheet
                        DReasonCode.Visible = False     '延遲理由代碼
                        DReason.Visible = False         '延遲理由
                        DReasonDesc.Visible = False     '延遲其他說明
                    End If
                End If

            End If
        Else
            Top = 1950
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
        End If
        '按鈕及超連結設值

        If wStep = 1 Then
            Top = 2000
        Else
            Top = 2050
        End If


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"
        DHistoryLabel.Style("top") = Top + 32 & "px"
        GridView2.Style("top") = Top + 32 + 16 & "px"
        '簽核,通知 移除控制設定

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


        Dim InputData1 As String
        Dim InputData2 As String
        Dim InputData3 As String

        Dim SQL As String

        InputData1 = D1.Text
        InputData2 = D2.Text
        InputData3 = DCODE.Text



        '帶入CUSTOMER
        If InputData1 <> "" Then

            SQL = " Select  * from  MST_Custmer where 1=1 "
            If InputData1 <> "" Then
                SQL = SQL + " and ( custmer like '%" + InputData1 + "%' or name_c like '%" + InputData1 + "%')"
            End If
            SQL = SQL + " order by custmer,name_c "

            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DCUSTOMER.Text = DBData.Rows(0).Item("Name_C")
            DCUSTOMERCODE.Text = DBData.Rows(0).Item("Custmer")
        End If

        '帶入BUYER
        If InputData2 <> "" Then


            SQL = " Select  * from  MST_Buyer where 1=1 "
            If InputData2 <> "" Then
                SQL = SQL + " and ( buyer like '%" + InputData2 + "%' or buyer_name like '%" + InputData2 + "%')"
            End If

            SQL = SQL + " order by buyer_name,buyer "
            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DBUYER.Text = DBData.Rows(0).Item("Buyer_Name")
            DBUYERCODE.Text = DBData.Rows(0).Item("Buyer")

        End If


        '帶入ITEM  '20210913
        If InputData3 <> "" Then


            SQL = "  select item,replace(replace(itemname,'<','('),'>',')')  itemname from (  Select item,rtrim(item_name1)+' '+rtrim(item_name2)+' '+ltrim(item_name3) as itemname from Mst_item where 1=1 "
            If InputData3 <> "" Then
                SQL = SQL + " and ( item ='" + InputData3 + "' ))a "
            End If

            SQL = SQL + " order by item "
            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DSPECNAME.Text = DBData.Rows(0).Item("itemname")
            DSPEC.Text = DBData.Rows(0).Item("item")

        End If





        'No
        Select Case FindFieldInf("NO")
            Case 0  '顯示
                DNO.BackColor = Color.White
                DNO.Visible = True
                DNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNO.Visible = True
                DNO.BackColor = Color.GreenYellow
                DNO.ReadOnly = False
                ShowRequiredFieldValidator("DNoRqd", "DNo", "異常：需輸入Ｎｏ")
            Case 2  '修改
                DNO.Visible = True
                DNO.BackColor = Color.Yellow
                DNO.ReadOnly = False
            Case Else   '隱藏
                DNO.Visible = False
        End Select

        'Date
        Select Case FindFieldInf("DATE")
            Case 0  '顯示
                DDATE.BackColor = Color.LightGray
                DDATE.Visible = True
                DDATE.Attributes.Add("readonly", "true")
                BDATE.Visible = False
            Case 1  '修改+檢查
                DDATE.Visible = True
                DDATE.BackColor = Color.GreenYellow
                DDATE.ReadOnly = False
                BDATE.Visible = True
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入申請日期")
            Case 2  '修改
                DDATE.Visible = True
                DDATE.BackColor = Color.Yellow
                DDATE.ReadOnly = False
                BDATE.Visible = True
            Case Else   '隱藏
                DDATE.Visible = False
                BDATE.Visible = False
        End Select
        If pPost = "New" Then DDATE.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '區域
        Select Case FindFieldInf("GLOBAL")
            Case 0  '顯示
                DGLOBAL.BackColor = Color.LightGray
                DGLOBAL.Visible = True

            Case 1  '修改+檢查
                DGLOBAL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DGLOBALRqd", "DGLOBAL", "異常：需輸入極/地域")
                DGLOBAL.Visible = True
            Case 2  '修改
                DGLOBAL.BackColor = Color.Yellow
                DGLOBAL.Visible = True
            Case Else   '隱藏
                DGLOBAL.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("GLOBAL", "ZZZZZZ")


        'Customer
        Select Case FindFieldInf("CUSTOMER")
            Case 0  '顯示
                DCUSTOMER.BackColor = Color.LightGray
                DCUSTOMER.ReadOnly = True
                DCUSTOMER.Visible = True
                DCUSTOMERCODE.BackColor = Color.LightGray
                DCUSTOMERCODE.ReadOnly = True
                DCUSTOMERCODE.Visible = True
                BCustomer.Visible = False

            Case 1  '修改+檢查
                DCUSTOMER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerRqd", "DCustomer", "異常：需輸入客戶名稱")
                DCUSTOMER.Visible = True
                DCUSTOMER.ReadOnly = True
                DCUSTOMERCODE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "異常：需輸入客戶名稱")
                DCUSTOMERCODE.Visible = True
                DCUSTOMERCODE.ReadOnly = True
                BCustomer.Visible = True

            Case 2  '修改
                DCUSTOMER.BackColor = Color.Yellow
                DCUSTOMER.Visible = True
                DCUSTOMER.ReadOnly = True
                DCUSTOMERCODE.BackColor = Color.Yellow
                DCUSTOMERCODE.Visible = True
                DCUSTOMERCODE.ReadOnly = True
                BCustomer.Visible = True

            Case Else   '隱藏
                DCUSTOMER.Visible = False
                DCUSTOMERCODE.Visible = False
                BCustomer.Visible = False

        End Select
        ' If pPost = "New" Then SetFieldData("CUSTOMER", "ZZZZZZ")



        '營業人員
        Select Case FindFieldInf("APPNAME")
            Case 0  '顯示
                DAPPNAME.BackColor = Color.LightGray
                DAPPNAME.Visible = True

            Case 1  '修改+檢查
                DAPPNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAPPNAMERqd", "DAPPNAME", "異常：需輸入營業人員")
                DAPPNAME.Visible = True
            Case 2  '修改
                DAPPNAME.BackColor = Color.Yellow
                DAPPNAME.Visible = True
            Case Else   '隱藏
                DAPPNAME.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("APPNAME", "ZZZZZZ")


        '樣品
        Select Case FindFieldInf("PLACE")
            Case 0  '顯示
                DSAMPLE.BackColor = Color.LightGray
                DSAMPLE.Visible = True

            Case 1  '修改+檢查
                DSAMPLE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSAMPLERqd", "DSAMPLE", "異常：需輸入樣品")
                DSAMPLE.Visible = True
            Case 2  '修改
                DSAMPLE.BackColor = Color.Yellow
                DSAMPLE.Visible = True
            Case Else   '隱藏
                DSAMPLE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SAMPLE", "ZZZZZZ")






        '原ORNO"
        Select Case FindFieldInf("OORNO")
            Case 0  '顯示
                DOORNO.BackColor = Color.LightGray
                DOORNO.Visible = True
                DOORNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DOORNO.Visible = True
                DOORNO.BackColor = Color.GreenYellow
                DOORNO.ReadOnly = False
                ShowRequiredFieldValidator("DOORNORqd", "DOORNO", "異常：需輸入原ORNO")
            Case 2  '修改
                DOORNO.Visible = True
                DOORNO.BackColor = Color.Yellow
                DOORNO.ReadOnly = False
            Case Else   '隱藏
                DOORNO.Visible = False
        End Select


        '新NOORNO"
        Select Case FindFieldInf("NORNO")
            Case 0  '顯示
                DNORNO.BackColor = Color.LightGray
                DNORNO.Visible = True
                DNORNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNORNO.Visible = True
                DNORNO.BackColor = Color.GreenYellow
                DNORNO.ReadOnly = False
                ShowRequiredFieldValidator("DNORNORqd", "DNORNO", "異常：需輸入新ORNO")
            Case 2  '修改
                DNORNO.Visible = True
                DNORNO.BackColor = Color.Yellow
                DNORNO.ReadOnly = False
            Case Else   '隱藏
                DNORNO.Visible = False
        End Select




        'SPEC
        Select Case FindFieldInf("SPEC")
            Case 0  '顯示
                DSPEC.BackColor = Color.LightGray
                DSPEC.ReadOnly = True
                DSPEC.Visible = True
                BBuyer.Visible = False
            Case 1  '修改+檢查
                DSPEC.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSPECRqd", "DSPEC", "異常：需輸入型別")
                DSPEC.Visible = True
                DSPEC.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '修改
                DSPEC.BackColor = Color.Yellow
                DSPEC.Visible = True
                BBuyer.Visible = True
                DSPEC.ReadOnly = True
            Case Else   '隱藏
                DSPEC.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SPEC", "ZZZZZZ")






        'SPECNAME
        Select Case FindFieldInf("SPEC")
            Case 0  '顯示
                DSPECNAME.BackColor = Color.LightGray
                DSPECNAME.ReadOnly = True
                DSPECNAME.Visible = True
                BSPEC.Visible = False
            Case 1  '修改+檢查
                DSPECNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSPECNAMERqd", "DSPECNAME", "異常：需輸入型別")
                DSPECNAME.Visible = True
                DSPECNAME.ReadOnly = True
                BSPEC.Visible = True
            Case 2  '修改
                DSPECNAME.BackColor = Color.Yellow
                DSPECNAME.Visible = True
                BBuyer.Visible = True
                DSPECNAME.ReadOnly = True
            Case Else   '隱藏
                DSPECNAME.Visible = False
                BSPEC.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SPECNAME", "ZZZZZZ")



        'OR數量
        Select Case FindFieldInf("ORQTY")
            Case 0  '顯示
                DORQTY.BackColor = Color.LightGray
                DORQTY.Visible = True
                DORQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DORQTY.Visible = True
                DORQTY.BackColor = Color.GreenYellow
                DORQTY.ReadOnly = False
                ShowRequiredFieldValidator("DORQTYRqd", "DORQTY", "異常：需輸入OR數量")
            Case 2  '修改
                DORQTY.Visible = True
                DORQTY.BackColor = Color.Yellow
                DORQTY.ReadOnly = False
            Case Else   '隱藏
                DORQTY.Visible = False
        End Select




        'SHIPDATE
        Select Case FindFieldInf("SHIPDATE")
            Case 0  '顯示
                DSHIPDATE.BackColor = Color.LightGray
                DSHIPDATE.Visible = True
                DSHIPDATE.Attributes.Add("readonly", "true")
                BSHIPDATE.Visible = False

                DORDERDATE.BackColor = Color.LightGray
                DORDERDATE.Visible = True
                DORDERDATE.Attributes.Add("readonly", "true")
                BORDERDATE.Visible = False



            Case 1  '修改+檢查
                DSHIPDATE.Visible = True
                DSHIPDATE.BackColor = Color.GreenYellow
                DSHIPDATE.ReadOnly = False
                BSHIPDATE.Visible = True
                ShowRequiredFieldValidator("DSHIPDATERqd", "DSHIPDATE", "異常：需輸入出貨日期")

                DORDERDATE.Visible = True
                DORDERDATE.BackColor = Color.GreenYellow
                DORDERDATE.ReadOnly = False
                BORDERDATE.Visible = True
                ShowRequiredFieldValidator("DORDERDATERqd", "DORDERDATE", "異常：需輸入訂貨日期")


            Case 2  '修改
                DSHIPDATE.Visible = True
                DSHIPDATE.BackColor = Color.Yellow
                DSHIPDATE.ReadOnly = False
                BSHIPDATE.Visible = True

                DORDERDATE.Visible = True
                DORDERDATE.BackColor = Color.Yellow
                DORDERDATE.ReadOnly = False
                BORDERDATE.Visible = True

            Case Else   '隱藏
                DSHIPDATE.Visible = False
                BSHIPDATE.Visible = False

                DORDERDATE.Visible = False
                BORDERDATE.Visible = False

        End Select
        '  If pPost = "New" Then DSHIPDATE.Text = Now.ToString("yyyy/MM/dd") '現在日時
        ' If pPost = "New" Then DORDERDATE.Text = Now.ToString("yyyy/MM/dd") '現在日時



        '大貨
        Select Case FindFieldInf("BIGGOODS")
            Case 0  '顯示
                DBIGGOODS.BackColor = Color.LightGray
                DBIGGOODS.Visible = True

            Case 1  '修改+檢查
                DBIGGOODS.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBIGGOODSRqd", "DBIGGOODS", "異常：需輸入大貨")
                DBIGGOODS.Visible = True
            Case 2  '修改
                DBIGGOODS.BackColor = Color.Yellow
                DBIGGOODS.Visible = True
            Case Else   '隱藏
                DBIGGOODS.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BIGGOODS", "ZZZZZZ")


        '地點
        Select Case FindFieldInf("PLACE")
            Case 0  '顯示
                DPLACE.BackColor = Color.LightGray
                DPLACE.Visible = True

                DSAMPLE.BackColor = Color.LightGray
                DSAMPLE.Visible = True


                DDELIVERYDATE.BackColor = Color.LightGray
                DDELIVERYDATE.Visible = True
                DDELIVERYDATE.Attributes.Add("readonly", "true")
                BDELIVERYDATE.Visible = False

            Case 1  '修改+檢查
                DPLACE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPLACERqd", "DPLACE", "異常：需輸入地點")
                DPLACE.Visible = True


                DSAMPLE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSAMPLERqd", "DSAMPLE", "異常：需輸入樣品")
                DSAMPLE.Visible = True


                DDELIVERYDATE.Visible = True
                DDELIVERYDATE.BackColor = Color.GreenYellow
                DDELIVERYDATE.ReadOnly = False
                BDELIVERYDATE.Visible = True
                ShowRequiredFieldValidator("DELIVERYDATERqd", "DORDERDATE", "異常：需輸入期望納期")

            Case 2  '修改
                DPLACE.BackColor = Color.Yellow
                DPLACE.Visible = True

                DSAMPLE.BackColor = Color.Yellow
                DSAMPLE.Visible = True

                DDELIVERYDATE.Visible = True
                DDELIVERYDATE.BackColor = Color.Yellow
                DDELIVERYDATE.ReadOnly = False
                BDELIVERYDATE.Visible = True


            Case Else   '隱藏
                DPLACE.Visible = False
                DSAMPLE.Visible = False

                DDELIVERYDATE.Visible = False
                BDELIVERYDATE.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("PLACE", "ZZZZZZ")
        If pPost = "New" Then SetFieldData("SAMPLE", "ZZZZZZ")



        '中文客訴內容
        Select Case FindFieldInf("CPCONTENT")
            Case 0  '顯示
                DCPCONTENT.BackColor = Color.LightGray
                DCPCONTENT.Visible = True
                DCPCONTENT.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCPCONTENT.Visible = True
                DCPCONTENT.BackColor = Color.GreenYellow
                DCPCONTENT.ReadOnly = False
                ShowRequiredFieldValidator("DCPCONTENTRqd", "DCPCONTENT", "異常：需輸入中文客訴內容")
            Case 2  '修改
                DCPCONTENT.Visible = True
                DCPCONTENT.BackColor = Color.Yellow
                DCPCONTENT.ReadOnly = False
            Case Else   '隱藏
                DCPCONTENT.Visible = False
        End Select





        '英文客訴內容
        Select Case FindFieldInf("CPCONTENT")
            Case 0  '顯示
                DEPCONTENT.BackColor = Color.LightGray
                DEPCONTENT.Visible = True
                DEPCONTENT.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEPCONTENT.Visible = True
                DEPCONTENT.BackColor = Color.GreenYellow
                DEPCONTENT.ReadOnly = False
                ShowRequiredFieldValidator("DEPCONTENTRqd", "DEPCONTENT", "異常：需輸入英文客訴內容")
            Case 2  '修改
                DEPCONTENT.Visible = True
                DEPCONTENT.BackColor = Color.Yellow
                DEPCONTENT.ReadOnly = False
            Case Else   '隱藏
                DEPCONTENT.Visible = False
        End Select



        '客訴數量
        Select Case FindFieldInf("CPQTY")
            Case 0  '顯示
                DCPQTY.BackColor = Color.LightGray
                DCPQTY.Visible = True
                DCPQTY.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCPQTY.Visible = True
                DCPQTY.BackColor = Color.GreenYellow
                DCPQTY.ReadOnly = False
                ShowRequiredFieldValidator("DCPQTYRqd", "DCPQTY", "異常：需輸入客訴數量")
            Case 2  '修改
                DCPQTY.Visible = True
                DCPQTY.BackColor = Color.Yellow
                DCPQTY.ReadOnly = False
            Case Else   '隱藏
                DCPQTY.Visible = False
        End Select



        'Buyer
        Select Case FindFieldInf("BUYER")
            Case 0  '顯示
                DBUYER.BackColor = Color.LightGray
                DBUYER.ReadOnly = True
                DBUYER.Visible = True
                BBuyer.Visible = False
            Case 1  '修改+檢查
                DBUYER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBUYER.Visible = True
                DBUYER.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '修改
                DBUYER.BackColor = Color.Yellow
                DBUYER.Visible = True
                BBuyer.Visible = True
                DBUYER.ReadOnly = True

            Case Else   '隱藏
                DBUYER.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BUYER", "ZZZZZZ")


        'BuyerCode
        Select Case FindFieldInf("BUYER")
            Case 0  '顯示
                DBUYERCODE.BackColor = Color.LightGray
                DBUYERCODE.ReadOnly = True
                DBUYERCODE.Visible = True
                BBuyer.Visible = False
            Case 1  '修改+檢查
                DBUYERCODE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerCodeRqd", "DBuyerCode", "異常：需輸入Buyer")
                DBUYERCODE.Visible = True
                DBUYERCODE.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '修改
                DBUYERCODE.BackColor = Color.Yellow
                DBUYERCODE.Visible = True
                BBuyer.Visible = True
                DBUYERCODE.ReadOnly = True
            Case Else   '隱藏
                DBUYERCODE.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BuyerCode", "ZZZZZZ")



        '簡圖
        Select Case FindFieldInf("MapFile")
            Case 0  '顯示
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "異常：需上傳不良樣品圖")
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


        '分野項目
        Select Case FindFieldInf("ITEM")
            Case 0  '顯示
                DITEM.BackColor = Color.LightGray
                DITEM.Visible = True

            Case 1  '修改+檢查
                DITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DITEMRqd", "DITEM", "異常：需輸入分野項目")
                DITEM.Visible = True
            Case 2  '修改
                DITEM.BackColor = Color.Yellow
                DITEM.Visible = True
            Case Else   '隱藏
                DITEM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ITEM", "ZZZZZZ")





        '顧客回覆書


        '分野項目
        Select Case FindFieldInf("REPLAYLAN")
            Case 0  '顯示
                DREPLAYLAN.BackColor = Color.LightGray
                DREPLAYLAN.Visible = True

            Case 1  '修改+檢查
                DREPLAYLAN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DREPLAYLANRqd", "DREPLAYLAN", "異常：需輸入顧客回覆書")
                DREPLAYLAN.Visible = True
            Case 2  '修改
                DREPLAYLAN.BackColor = Color.Yellow
                DREPLAYLAN.Visible = True
            Case Else   '隱藏
                DREPLAYLAN.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("REPLAYLAN", "ZZZZZZ")


        '********************************************************************************************************************************************************
        ' 品管
        '********************************************************************************************************************************************************


        '品管受理編號
        Select Case FindFieldInf("QCNO")
            Case 0  '顯示
                DQCNO.BackColor = Color.LightGray
                DQCNO.Visible = True
                DQCNO.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCNO.Visible = True
                DQCNO.BackColor = Color.GreenYellow
                DQCNO.ReadOnly = False
                ' ShowRequiredFieldValidator("DQCNORqd", "DQCNO", "異常：需輸入品管受理編號")


            Case 2  '修改
                DQCNO.Visible = True
                DQCNO.BackColor = Color.Yellow
                DQCNO.ReadOnly = False

            Case Else   '隱藏
                DQCNO.Visible = False

        End Select


        '責任部門別
        Select Case FindFieldInf("ACCDEP1")
            Case 0  '顯示
                DACCDEP1.BackColor = Color.LightGray
                DACCDEP1.Visible = True
                DACCDEP12.BackColor = Color.LightGray
                DACCDEP12.Visible = True
                DACCDEP13.BackColor = Color.LightGray
                DACCDEP13.Visible = True

            Case 1  '修改+檢查
                DACCDEP1.BackColor = Color.GreenYellow
                '   ShowRequiredFieldValidator("DACCDEP1Rqd", "DACCDEP1", "異常：需輸入責任部門別")
                DACCDEP1.Visible = True

                DACCDEP12.BackColor = Color.GreenYellow
                '   ShowRequiredFieldValidator("DACCDEP1Rqd", "DACCDEP1", "異常：需輸入責任部門別")
                DACCDEP12.Visible = True

                DACCDEP13.BackColor = Color.GreenYellow
                '   ShowRequiredFieldValidator("DACCDEP1Rqd", "DACCDEP1", "異常：需輸入責任部門別")
                DACCDEP13.Visible = True

            Case 2  '修改
                DACCDEP1.BackColor = Color.Yellow
                DACCDEP1.Visible = True
                DACCDEP12.BackColor = Color.Yellow
                DACCDEP12.Visible = True
                DACCDEP13.BackColor = Color.Yellow
                DACCDEP13.Visible = True

            Case Else   '隱藏
                DACCDEP1.Visible = False
                DACCDEP12.Visible = False
                DACCDEP13.Visible = False

        End Select


        If pPost = "New" Then SetFieldData("ACCDEP1", "ZZZZZZ")
        If pPost = "New" Then SetFieldData("ACCDEP12", "ZZZZZZ")
        If pPost = "New" Then SetFieldData("ACCDEP13", "ZZZZZZ")



        '客訴判定
        Select Case FindFieldInf("TYPE")
            Case 0  '顯示
                DTYPE.BackColor = Color.LightGray
                DTYPE.Visible = True
                DTYPE1.BackColor = Color.LightGray
                DTYPE1.Visible = True
                DTYPE2.BackColor = Color.LightGray
                DTYPE2.Visible = True

                DQCDESCTYPE.BackColor = Color.LightGray
                DQCDESCTYPE.Visible = True

            Case 1  '修改+檢查
                DTYPE.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DTYPERqd", "DTYPE", "異常：需輸入客訴判定")
                DTYPE.Visible = True

                DTYPE1.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DTYPERqd", "DTYPE", "異常：需輸入客訴判定")
                DTYPE1.Visible = True

                DTYPE2.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DTYPERqd", "DTYPE", "異常：需輸入客訴判定")
                DTYPE2.Visible = True

                DQCDESCTYPE.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DTYPERqd", "DTYPE", "異常：需輸入客訴判定")
                DQCDESCTYPE.Visible = True

            Case 2  '修改
                DTYPE.BackColor = Color.Yellow
                DTYPE.Visible = True

                DTYPE1.BackColor = Color.Yellow
                DTYPE1.Visible = True

                DTYPE2.BackColor = Color.Yellow
                DTYPE2.Visible = True

                DQCDESCTYPE.BackColor = Color.Yellow
                DQCDESCTYPE.Visible = True

            Case Else   '隱藏
                DTYPE.Visible = False
                DTYPE1.Visible = False
                DTYPE2.Visible = False

                DQCDESCTYPE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("TYPE", "ZZZZZZ")
        If pPost = "New" Then SetFieldData("TYPE2", "ZZZZZZ")
        If pPost = "New" Then SetFieldData("QCDESCTYPE", "ZZZZZZ")

        'PL
        Select Case FindFieldInf("TYPE")
            Case 0  '顯示
                DPL.BackColor = Color.LightGray
                DPL.Visible = True


            Case 1  '修改+檢查
                DPL.BackColor = Color.GreenYellow
                'ShowRequiredFieldValidator("DPLRqd", "DPL", "異常：需輸入PL")
                DPL.Visible = True


            Case 2  '修改
                DPL.BackColor = Color.Yellow
                DPL.Visible = True

            Case Else   '隱藏
                DPL.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("PL", "ZZZZZZ")





        '責任部門別
        Select Case FindFieldInf("ACCDEP2")
            Case 0  '顯示

                DACCDEP2.BackColor = Color.LightGray
                DACCDEP2.Visible = True
                DACCDEP2NO.BackColor = Color.LightGray
                DACCDEP2NO.ReadOnly = True

            Case 1  '修改+檢查

                DACCDEP2.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DACCDEP2Rqd", "DACCDEP2", "異常：需輸入責任工程別")
                DACCDEP2.Visible = True
                DACCDEP2NO.BackColor = Color.LightGray
                DACCDEP2NO.ReadOnly = True
            Case 2  '修改
                DACCDEP2.BackColor = Color.Yellow
                DACCDEP2.Visible = True
                DACCDEP2NO.BackColor = Color.LightGray
                DACCDEP2NO.ReadOnly = True
            Case Else   '隱藏

                DACCDEP2.Visible = False
                DACCDEP2NO.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("ACCDEP2", "ZZZZZZ")


        '責任擔當者
        Select Case FindFieldInf("ACCEMPNAME")
            Case 0  '顯示            
                DACCEMPNAME.BackColor = Color.LightGray
                DACCEMPNAME.Visible = True
                DACCEMPNAME.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DACCEMPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.GreenYellow
                DACCEMPNAME.ReadOnly = False
                '  ShowRequiredFieldValidator("DACCEMPNAMERqd", "DACCEMPNAME", "異常：需輸入責任擔當者")

            Case 2  '修改
                DACCEMPNAME.Visible = True
                DACCEMPNAME.BackColor = Color.Yellow
                DACCEMPNAME.ReadOnly = False

            Case Else   '隱藏
                DACCEMPNAME.Visible = False
        End Select



        '********************************************************************************************************************************************************
        ' 製造評鑑說明
        '********************************************************************************************************************************************************

        '顧客 
        Select Case FindFieldInf("CUSTOMERTYPE")
            Case 0  '顯示
                DCUSTOMERTYPE.BackColor = Color.LightGray
                DCUSTOMERTYPE.Visible = True


            Case 1  '修改+檢查
                DCUSTOMERTYPE.BackColor = Color.GreenYellow
                '        ShowRequiredFieldValidator("DCUSTOMERTYPERqd", "DCUSTOMERTYPE", "異常：需輸入顧客")
                DCUSTOMERTYPE.Visible = True


            Case 2  '修改
                DCUSTOMERTYPE.BackColor = Color.Yellow
                DCUSTOMERTYPE.Visible = True

            Case Else   '隱藏
                DCUSTOMERTYPE.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("CUSTOMERTYPE", "ZZZZZZ")



        '判定
        Select Case FindFieldInf("JUDGE")
            Case 0  '顯示
                DJUDGE.BackColor = Color.LightGray
                DJUDGE.Visible = True
                DJUDGE1.BackColor = Color.LightGray
                DJUDGE1.Visible = True
                DJUDGE2.BackColor = Color.LightGray
                DJUDGE2.Visible = True
            Case 1  '修改+檢查
                DJUDGE.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DJUDGERqd", "DJUDGE", "異常：需輸入判定")
                DJUDGE.Visible = True

                DJUDGE1.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DJUDGERqd", "DJUDGE", "異常：需輸入判定")
                DJUDGE1.Visible = True

                DJUDGE2.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DJUDGERqd", "DJUDGE", "異常：需輸入判定")
                DJUDGE2.Visible = True


            Case 2  '修改
                DJUDGE.BackColor = Color.Yellow
                DJUDGE.Visible = True

                DJUDGE1.BackColor = Color.Yellow
                DJUDGE1.Visible = True

                DJUDGE2.BackColor = Color.Yellow
                DJUDGE2.Visible = True
            Case Else   '隱藏
                DJUDGE.Visible = False
                DJUDGE1.Visible = False
                DJUDGE2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("JUDGE", "ZZZZZZ")


        '原因人為錯誤分類
        Select Case FindFieldInf("MANUREASON")
            Case 0  '顯示
                DMANUREASON.BackColor = Color.LightGray
                DMANUREASON.Visible = True
                DMANUREASON1.BackColor = Color.LightGray
                DMANUREASON1.Visible = True


            Case 1  '修改+檢查
                DMANUREASON.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DMANUREASONRqd", "DMANUREASON", "異常：需輸入原因")
                DMANUREASON.Visible = True

                '   DMANUREASON1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMANUREASON1Rqd", "DMANUREASON1", "異常：需輸入人為因素")
                DMANUREASON1.Visible = True

            Case 2  '修改
                DMANUREASON.BackColor = Color.Yellow
                DMANUREASON.Visible = True

                DMANUREASON1.BackColor = Color.Yellow
                DMANUREASON1.Visible = True

            Case Else   '隱藏
                DMANUREASON.Visible = False
                DMANUREASON1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MANUREASON", "ZZZZZZ")
        If pPost = "New" Then SetFieldData("MANUREASON1", "ZZZZZZ")




        '責任區分起因
        Select Case FindFieldInf("RESPONS")
            Case 0  '顯示
                DRESPONS.BackColor = Color.LightGray
                DRESPONS.Visible = True


            Case 1  '修改+檢查
                DRESPONS.BackColor = Color.GreenYellow
                '   ShowRequiredFieldValidator("DRESPONSRqd", "DRESPONS", "異常：需輸入責任區分起因")
                DRESPONS.Visible = True


            Case 2  '修改
                DRESPONS.BackColor = Color.Yellow
                DRESPONS.Visible = True

            Case Else   '隱藏
                DRESPONS.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("RESPONS", "ZZZZZZ")



        '檢查工時
        Select Case FindFieldInf("HOUR1")
            Case 0  '顯示            
                DHOUR1.BackColor = Color.LightGray
                DHOUR1.Visible = True
                DHOUR1.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DHOUR1.Visible = True
                DHOUR1.BackColor = Color.GreenYellow
                DHOUR1.ReadOnly = False
                ' ShowRequiredFieldValidator("DHOUR1Rqd", "DHOUR1", "異常：需輸入處理工時")

            Case 2  '修改
                DHOUR1.Visible = True
                DHOUR1.BackColor = Color.Yellow
                DHOUR1.ReadOnly = False

            Case Else   '隱藏
                DHOUR1.Visible = False
        End Select
        '檢查工時

        Select Case FindFieldInf("HOUR2")
            Case 0  '顯示            
                DHOUR2.BackColor = Color.LightGray
                DHOUR2.Visible = True
                DHOUR2.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DHOUR2.Visible = True
                DHOUR2.BackColor = Color.GreenYellow
                DHOUR2.ReadOnly = False
                'ShowRequiredFieldValidator("DHOUR2Rqd", "DHOUR2", "異常：需輸入檢查工時")

            Case 2  '修改
                DHOUR2.Visible = True
                DHOUR2.BackColor = Color.Yellow
                DHOUR2.ReadOnly = False

            Case Else   '隱藏
                DHOUR2.Visible = False
        End Select

        Select Case FindFieldInf("HOUR3")
            Case 0  '顯示            
                DHOUR3.BackColor = Color.LightGray
                DHOUR3.Visible = True
                DHOUR3.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DHOUR3.Visible = True
                DHOUR3.BackColor = Color.GreenYellow
                DHOUR3.ReadOnly = False
                'ShowRequiredFieldValidator("DHOUR3Rqd", "DHOUR3", "異常：需輸入國內出差工時")

            Case 2  '修改
                DHOUR3.Visible = True
                DHOUR3.BackColor = Color.Yellow
                DHOUR3.ReadOnly = False

            Case Else   '隱藏
                DHOUR3.Visible = False
        End Select


        Select Case FindFieldInf("HOUR4")
            Case 0  '顯示            
                DHOUR4.BackColor = Color.LightGray
                DHOUR4.Visible = True
                DHOUR4.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DHOUR4.Visible = True
                DHOUR4.BackColor = Color.GreenYellow
                DHOUR4.ReadOnly = False
                'ShowRequiredFieldValidator("DHOUR4Rqd", "DHOUR4", "異常：需輸入現場檢查工時")

            Case 2  '修改
                DHOUR4.Visible = True
                DHOUR4.BackColor = Color.Yellow
                DHOUR4.ReadOnly = False

            Case Else   '隱藏
                DHOUR4.Visible = False
        End Select



        Select Case FindFieldInf("HOUR5")
            Case 0  '顯示            
                DHOUR5.BackColor = Color.LightGray
                DHOUR5.Visible = True
                DHOUR5.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DHOUR5.Visible = True
                DHOUR5.BackColor = Color.GreenYellow
                DHOUR5.ReadOnly = False
                'ShowRequiredFieldValidator("DHOUR4Rqd", "DHOUR4", "異常：需輸入其它費用")

            Case 2  '修改
                DHOUR5.Visible = True
                DHOUR5.BackColor = Color.Yellow
                DHOUR5.ReadOnly = False

            Case Else   '隱藏
                DHOUR5.Visible = False
        End Select


        'MFGDATE
        Select Case FindFieldInf("JUDGE")
            Case 0  '顯示
                DMFGDate.BackColor = Color.LightGray
                DMFGDate.Visible = True
                DMFGDate.Attributes.Add("readonly", "true")
                BMFGDate.Visible = False
            Case 1  '修改+檢查
                DMFGDate.Visible = True
                DMFGDate.BackColor = Color.GreenYellow
                DMFGDate.ReadOnly = False
                BMFGDate.Visible = True
                ShowRequiredFieldValidator("DMFGDATERqd", "DMFGDATE", "異常：需輸入生產日期")
             
            Case 2  '修改
                DMFGDate.Visible = True
                DMFGDate.BackColor = Color.Yellow
                DMFGDate.ReadOnly = False
                BMFGDate.Visible = True
            Case Else   '隱藏
                DMFGDate.Visible = False
                BMFGDate.Visible = False
        End Select
        '        If pPost = "New" Then DMFGDate.Text = Now.ToString("yyyy/MM/dd") '現在日時



        '********************************************************************************************************************************************************
        ' 品管回覆
        '********************************************************************************************************************************************************

        '收到樣品日

        Select Case FindFieldInf("SAMPLEDATE")
            Case 0  '顯示
                DSAMPLEDATE.BackColor = Color.LightGray
                DSAMPLEDATE.Visible = True
                DSAMPLEDATE.Attributes.Add("readonly", "true")
                BSAMPLEDATE.Visible = False

                DCARDATE.BackColor = Color.LightGray
                DCARDATE.Visible = True
                DCARDATE.Attributes.Add("readonly", "true")
                BCARDATE.Visible = False

                DCARDATE.BackColor = Color.LightGray
                DCARDATE.Visible = True
                DCARDATE.Attributes.Add("readonly", "true")
                BCARDATE.Visible = False


                DFIRSTDATE.BackColor = Color.LightGray
                DFIRSTDATE.Visible = True
                DFIRSTDATE.Attributes.Add("readonly", "true")
                BFIRSTDATE.Visible = False

                DFINALDATE.BackColor = Color.LightGray
                DFINALDATE.Visible = True
                DFINALDATE.Attributes.Add("readonly", "true")
                BFINALDATE.Visible = False

                DMIDDATE.BackColor = Color.LightGray
                DMIDDATE.Visible = True
                DMIDDATE.Attributes.Add("readonly", "true")
                BMIDDATE.Visible = False


            Case 1  '修改+檢查
                DSAMPLEDATE.Visible = True
                DSAMPLEDATE.BackColor = Color.GreenYellow
                DSAMPLEDATE.ReadOnly = False
                BSAMPLEDATE.Visible = True
                ShowRequiredFieldValidator("DSAMPLEDATERqd", "DSAMPLEDATE", "異常：需輸入收到樣品日")

                DCARDATE.Visible = True
                DCARDATE.BackColor = Color.GreenYellow
                DCARDATE.ReadOnly = False
                BCARDATE.Visible = True
                ShowRequiredFieldValidator("DCARDATERqd", "DCARDATE", "異常：需輸入CAR收到日")


                DFIRSTDATE.Visible = True
                DFIRSTDATE.BackColor = Color.GreenYellow
                DFIRSTDATE.ReadOnly = False
                BFIRSTDATE.Visible = True
                ShowRequiredFieldValidator("DFIRSTDATERqd", "DFIRSTDATE", "異常：需輸入廠長初次報告日")

                DFINALDATE.Visible = True
                DFINALDATE.BackColor = Color.GreenYellow
                DFINALDATE.ReadOnly = False
                BFINALDATE.Visible = True
                ShowRequiredFieldValidator("DFINALDATERqd", "DFIRSTDATE", "異常：需輸入廠長最終報告日")

                DMIDDATE.Visible = True
                DMIDDATE.BackColor = Color.GreenYellow
                DMIDDATE.ReadOnly = False
                BMIDDATE.Visible = True
                ShowRequiredFieldValidator("DMIDDATERqd", "DMIDDATE", "異常：需輸入期中報告回覆日")



            Case 2  '修改
                DSAMPLEDATE.Visible = True
                DSAMPLEDATE.BackColor = Color.Yellow
                DSAMPLEDATE.ReadOnly = False
                BSAMPLEDATE.Visible = True

                DCARDATE.Visible = True
                DCARDATE.BackColor = Color.Yellow
                DCARDATE.ReadOnly = False
                BCARDATE.Visible = True


                DFIRSTDATE.Visible = True
                DFIRSTDATE.BackColor = Color.Yellow
                DFIRSTDATE.ReadOnly = False
                BFIRSTDATE.Visible = True


                DFINALDATE.Visible = True
                DFINALDATE.BackColor = Color.Yellow
                DFINALDATE.ReadOnly = False
                BFINALDATE.Visible = True

                DMIDDATE.Visible = True
                DMIDDATE.BackColor = Color.Yellow
                DMIDDATE.ReadOnly = False
                BMIDDATE.Visible = True

            Case Else   '隱藏
                DSAMPLEDATE.Visible = False
                BSAMPLEDATE.Visible = False
                DCARDATE.Visible = False
                BCARDATE.Visible = False

                DFIRSTDATE.Visible = False
                BFIRSTDATE.Visible = False

                DFINALDATE.Visible = False
                BFINALDATE.Visible = False

                DMIDDATE.Visible = False
                BMIDDATE.Visible = False
        End Select


        '品管覆函日

        Select Case FindFieldInf("QCDATE")
            Case 0  '顯示
                DQCDATE.BackColor = Color.LightGray
                DQCDATE.Visible = True
                DQCDATE.Attributes.Add("readonly", "true")
                BQCDATE.Visible = False
            Case 1  '修改+檢查
                DQCDATE.Visible = True
                DQCDATE.BackColor = Color.GreenYellow
                DQCDATE.ReadOnly = False
                BQCDATE.Visible = True
                ShowRequiredFieldValidator("DQCDATERqd", "DQCDATE", "異常：需輸入品管覆函日")
            Case 2  '修改
                DQCDATE.Visible = True
                DQCDATE.BackColor = Color.Yellow
                DQCDATE.ReadOnly = False
                BQCDATE.Visible = True
            Case Else   '隱藏
                DQCDATE.Visible = False
                BQCDATE.Visible = False
        End Select


        '正式回答日數
        Select Case FindFieldInf("ANSWERDAYS")
            Case 0  '顯示            
                DANSWERDAYS.BackColor = Color.LightGray
                DANSWERDAYS.Visible = True
                DANSWERDAYS.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DANSWERDAYS.Visible = True
                DANSWERDAYS.BackColor = Color.GreenYellow
                DANSWERDAYS.ReadOnly = False
                ShowRequiredFieldValidator("DANSWERDAYSRqd", "DANSWERDAYS", "異常：需輸入正式回答日數")

            Case 2  '修改
                DANSWERDAYS.Visible = True
                DANSWERDAYS.BackColor = Color.Yellow
                DANSWERDAYS.ReadOnly = False

            Case Else   '隱藏
                DANSWERDAYS.Visible = False
        End Select



        '備註
        Select Case FindFieldInf("REMARK")
            Case 0  '顯示
                DREMARK.BackColor = Color.LightGray
                DREMARK.Visible = True


            Case 1  '修改+檢查
                DREMARK.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DREMARKRqd", "DREMARK", "異常：需輸入客訴判定分類")
                DREMARK.Visible = True


            Case 2  '修改
                DREMARK.BackColor = Color.Yellow
                DREMARK.Visible = True

            Case Else   '隱藏
                DREMARK.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("REMARK", "ZZZZZZ")



        Select Case FindFieldInf("ACOST")
            Case 0  '顯示            
                DACOST.BackColor = Color.LightGray
                DACOST.Visible = True
                DACOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DACOST.Visible = True
                DACOST.BackColor = Color.GreenYellow
                DACOST.ReadOnly = False
                ' ShowRequiredFieldValidator("DACOSTRqd", "DACOST", "異常：需輸入賠償金額")

            Case 2  '修改
                DACOST.Visible = True
                DACOST.BackColor = Color.Yellow
                DACOST.ReadOnly = False

            Case Else   '隱藏
                DACOST.Visible = False
        End Select



        Select Case FindFieldInf("BCOST")
            Case 0  '顯示            
                DBCOST.BackColor = Color.LightGray
                DBCOST.Visible = True
                DBCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DBCOST.Visible = True
                DBCOST.BackColor = Color.GreenYellow
                DBCOST.ReadOnly = False
                ' ShowRequiredFieldValidator("DBCOSTRqd", "DBCOST", "異常：需輸入出差費用")

            Case 2  '修改
                DBCOST.Visible = True
                DBCOST.BackColor = Color.Yellow
                DBCOST.ReadOnly = False

            Case Else   '隱藏
                DBCOST.Visible = False
        End Select



        Select Case FindFieldInf("CCOST")
            Case 0  '顯示            
                DCCOST.BackColor = Color.LightGray
                DCCOST.Visible = True
                DCCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DCCOST.Visible = True
                DCCOST.BackColor = Color.GreenYellow
                DCCOST.ReadOnly = False
                '  ShowRequiredFieldValidator("DCCOSTRqd", "DCCOST", "異常：需輸入處理費用")

            Case 2  '修改
                DCCOST.Visible = True
                DCCOST.BackColor = Color.Yellow
                DCCOST.ReadOnly = False

            Case Else   '隱藏
                DCCOST.Visible = False
        End Select



        Select Case FindFieldInf("DCOST")
            Case 0  '顯示            
                DDCOST.BackColor = Color.LightGray
                DDCOST.Visible = True
                DDCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DDCOST.Visible = True
                DDCOST.BackColor = Color.GreenYellow
                DDCOST.ReadOnly = False
                ShowRequiredFieldValidator("DCOSTRqd", "DCOST", "異常：需輸入檢查費用")

            Case 2  '修改
                DDCOST.Visible = True
                DDCOST.BackColor = Color.Yellow
                DDCOST.ReadOnly = False

            Case Else   '隱藏
                DDCOST.Visible = False
        End Select


        Select Case FindFieldInf("ECOST")
            Case 0  '顯示            
                DECOST.BackColor = Color.LightGray
                DECOST.Visible = True
                DECOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DECOST.Visible = True
                DECOST.BackColor = Color.GreenYellow
                DECOST.ReadOnly = False
                '  ShowRequiredFieldValidator("DECOSTRqd", "DECOST", "異常：需輸入其它費用")

            Case 2  '修改
                DECOST.Visible = True
                DECOST.BackColor = Color.Yellow
                DECOST.ReadOnly = False

            Case Else   '隱藏
                DECOST.Visible = False
        End Select


        Select Case FindFieldInf("ECOST")
            Case 0  '顯示            
                DFCOST.BackColor = Color.LightGray
                DFCOST.Visible = True
                DFCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DFCOST.Visible = True
                DFCOST.BackColor = Color.GreenYellow
                DFCOST.ReadOnly = False
                '  ShowRequiredFieldValidator("DFCOSTRqd", "DFCOST", "異常：需輸入其它費用")

            Case 2  '修改
                DFCOST.Visible = True
                DFCOST.BackColor = Color.Yellow
                DFCOST.ReadOnly = False

            Case Else   '隱藏
                DFCOST.Visible = False
        End Select



        Select Case FindFieldInf("ECOST")
            Case 0  '顯示            
                DGCOST.BackColor = Color.LightGray
                DGCOST.Visible = True
                DGCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DGCOST.Visible = True
                DGCOST.BackColor = Color.GreenYellow
                DGCOST.ReadOnly = False
                '  ShowRequiredFieldValidator("DGCOSTRqd", "DGCOST", "異常：需輸入其它費用")

            Case 2  '修改
                DGCOST.Visible = True
                DGCOST.BackColor = Color.Yellow
                DGCOST.ReadOnly = False

            Case Else   '隱藏
                DGCOST.Visible = False
        End Select



        Select Case FindFieldInf("ECOST")
            Case 0  '顯示            
                DHCOST.BackColor = Color.LightGray
                DHCOST.Visible = True
                DHCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DHCOST.Visible = True
                DHCOST.BackColor = Color.GreenYellow
                DHCOST.ReadOnly = False
                '  ShowRequiredFieldValidator("DHCOSTRqd", "DHCOST", "異常：需輸入其它費用")

            Case 2  '修改
                DHCOST.Visible = True
                DHCOST.BackColor = Color.Yellow
                DHCOST.ReadOnly = False

            Case Else   '隱藏
                DHCOST.Visible = False
        End Select




        Select Case FindFieldInf("ALLCOST")
            Case 0  '顯示            
                DALLCOST.BackColor = Color.LightGray
                DALLCOST.Visible = True
                DALLCOST.Attributes.Add("readonly", "true")

            Case 1  '修改+檢查
                DALLCOST.Visible = True
                DALLCOST.BackColor = Color.GreenYellow
                DALLCOST.ReadOnly = False
                ShowRequiredFieldValidator("DALLCOSTRqd", "DALLCOST", "異常：需輸入合計")

            Case 2  '修改
                DALLCOST.Visible = True
                DALLCOST.BackColor = Color.Yellow
                DALLCOST.ReadOnly = False

            Case Else   '隱藏
                DALLCOST.Visible = False
        End Select


        '1.重要度

        Select Case FindFieldInf("POINT1")
            Case 0  '顯示              
                DPOINT1.BackColor = Color.LightGray
                DPOINT1.Visible = True
                DPOINT1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DPOINT1.Visible = True
                DPOINT1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPOINT1Rqd", "DPOINT1", "異常：需輸入1.重要度")
            Case 2  '修改
                DPOINT1.Visible = True
                DPOINT1.BackColor = Color.Yellow

            Case Else   '隱藏
                DPOINT1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("POINT1", "ZZZZZZ")

        '1.重要度

        Select Case FindFieldInf("POINT2")
            Case 0  '顯示              
                DPOINT2.BackColor = Color.LightGray
                DPOINT2.Visible = True
                DPOINT2.Attributes.Add("readonly", "true")
                DPOINT3.BackColor = Color.LightGray
                DPOINT3.Visible = True
                DPOINT3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DPOINT2.Visible = True
                DPOINT2.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DPOINT2Rqd", "DPOINT2", "異常：需輸入2.加算點")
                DPOINT3.Visible = True
                DPOINT3.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DPOINT3Rqd", "DPOINT3", "異常：需輸入3.重要點合計（=1+2)")
            Case 2  '修改
                DPOINT2.Visible = True
                DPOINT2.BackColor = Color.Yellow
                DPOINT3.Visible = True
                DPOINT3.BackColor = Color.Yellow
            Case Else   '隱藏
                DPOINT2.Visible = False
                DPOINT3.Visible = False
        End Select

        If wStep < 35 Or wStep = 500 Then
            DPOINT3.Visible = False
        Else

        End If





        ''1.重要度
        'Select Case FindFieldInf("POINT3")
        '    Case 0  '顯示              
        '        DPOINT3.BackColor = Color.LightGray
        '        DPOINT3.Visible = True
        '        DPOINT3.Attributes.Add("readonly", "true")
        '    Case 1  '修改+檢查
        '        DPOINT3.Visible = True
        '        DPOINT3.BackColor = Color.GreenYellow

        '        ShowRequiredFieldValidator("DPOINT3Rqd", "DPOINT3", "異常：需輸入3.重要點合計（=1+2)")
        '    Case 2  '修改
        '        DPOINT3.Visible = True
        '        DPOINT3.BackColor = Color.Yellow

        '    Case Else   '隱藏
        '        DPOINT3.Visible = False
        'End Select





        '1.全體評價
        Select Case FindFieldInf("SCORE")
            Case 0  '顯示              
                DSCORE1.BackColor = Color.LightGray
                DSCORE1.Visible = True
                DSCORE1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSCORE1.Visible = True
                DSCORE1.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DSCORE1Rqd", "DSCORE1", "異常：需輸入全體評價滿意度")
            Case 2  '修改
                DSCORE1.Visible = True
                DSCORE1.BackColor = Color.Yellow

            Case Else   '隱藏
                DSCORE1.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("SCORE1", "ZZZZZZ")


        '商品滿意度
        Select Case FindFieldInf("SCORE")
            Case 0  '顯示              
                DSCORE2.BackColor = Color.LightGray
                DSCORE2.Visible = True
                DSCORE2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSCORE2.Visible = True
                DSCORE2.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DSCORE2Rqd", "DSCORE2", "異常：需輸入商品滿意度")
            Case 2  '修改
                DSCORE2.Visible = True
                DSCORE2.BackColor = Color.Yellow

            Case Else   '隱藏
                DSCORE2.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("SCORE2", "ZZZZZZ")


        '服務滿意度
        Select Case FindFieldInf("SCORE")
            Case 0  '顯示              
                DSCORE3.BackColor = Color.LightGray
                DSCORE3.Visible = True
                DSCORE3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSCORE3.Visible = True
                DSCORE3.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DSCORE3Rqd", "DSCORE3", "異常：需輸入服務滿意度")
            Case 2  '修改
                DSCORE3.Visible = True
                DSCORE3.BackColor = Color.Yellow

            Case Else   '隱藏
                DSCORE3.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("SCORE3", "ZZZZZZ")


        '服務滿意度
        Select Case FindFieldInf("ACCDEP2")
            Case 0  '顯示              
                DHappen.BackColor = Color.LightGray
                DHappen.Visible = True
                DHappen.Attributes.Add("readonly", "true")
                DMach1.Visible = True
                DMach2.Visible = True
                DMACHNO.Visible = True
                DMACHNO.BackColor = Color.LightGray
            Case 1  '修改+檢查
                DHappen.Visible = True
                DHappen.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DHappenRqd", "DHappen", "異常：需輸入初次發生或再次發生")
                DMach1.Visible = True
                DMach2.Visible = True
                DMACHNO.Visible = True
                DMACHNO.BackColor = Color.GreenYellow
            Case 2  '修改
                DHappen.Visible = True
                DHappen.BackColor = Color.Yellow
                DMach1.Visible = True
                DMach2.Visible = True
                DMACHNO.Visible = True
                DMACHNO.BackColor = Color.Yellow
            Case Else   '隱藏
                DHappen.Visible = False
                DMach1.Visible = False
                DMach2.Visible = False
                DMACHNO.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("HAPPEN", "ZZZZZZ")




        '1.重要度
        Select Case FindFieldInf("COST")
            Case 0  '顯示              
                DCOST.BackColor = Color.LightGray
                DCOST.Visible = True
                DCOST.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCOST.Visible = True
                DCOST.BackColor = Color.GreenYellow
                DCOST.ReadOnly = False
                ShowRequiredFieldValidator("DCOSTRqd", "DCOST", "異常：需輸入費用分攤說明O")
            Case 2  '修改
                DCOST.Visible = True
                DCOST.BackColor = Color.Yellow
                DCOST.ReadOnly = False
            Case Else   '隱藏
                DCOST.Visible = False
        End Select




        '********************************************************************************************************************************************************
        ' 顧客回覆
        '********************************************************************************************************************************************************


        '顧客說明日

        Select Case FindFieldInf("CUSTDATE")
            Case 0  '顯示
                DCUSTDATE.BackColor = Color.LightGray
                DCUSTDATE.Visible = True
                DCUSTDATE.Attributes.Add("readonly", "true")
                BCUSTDATE.Visible = False
            Case 1  '修改+檢查
                DCUSTDATE.Visible = True
                DCUSTDATE.BackColor = Color.GreenYellow
                DCUSTDATE.ReadOnly = False
                BCUSTDATE.Visible = True
                '  ShowRequiredFieldValidator("DCUSTDATERqd", "DCUSTDATE", "異常：需輸入顧客說明日")
            Case 2  '修改
                DCUSTDATE.Visible = True
                DCUSTDATE.BackColor = Color.Yellow
                DCUSTDATE.ReadOnly = False
                BCUSTDATE.Visible = True
            Case Else   '隱藏
                DCUSTDATE.Visible = False
                BCUSTDATE.Visible = False
        End Select

        '顧客說明日

        Select Case FindFieldInf("YFGCDATE")
            Case 0  '顯示
                DYFGCDATE.BackColor = Color.LightGray
                DYFGCDATE.Visible = True
                DYFGCDATE.Attributes.Add("readonly", "true")
                BYFGCDATE.Visible = False
            Case 1  '修改+檢查
                DYFGCDATE.Visible = True
                DYFGCDATE.BackColor = Color.GreenYellow
                DYFGCDATE.ReadOnly = False
                BYFGCDATE.Visible = True
                '    ShowRequiredFieldValidator("DYFGCDATERqd", "DYFGCDATE", "異常：需輸入YFGC")
            Case 2  '修改
                DYFGCDATE.Visible = True
                DYFGCDATE.BackColor = Color.Yellow
                DYFGCDATE.ReadOnly = False
                BYFGCDATE.Visible = True
            Case Else   '隱藏
                DYFGCDATE.Visible = False
                BYFGCDATE.Visible = False
        End Select


        '
        Select Case FindFieldInf("Attachfile1")
            Case 0  '顯示            

                DAttachfile1.Visible = True
                DChkData1.Visible = True

            Case 1  '修改+檢查
                DAttachfile1.Visible = True
                DChkData1.Visible = True
            Case 2  '修改
                DAttachfile1.Visible = True
                DChkData1.Visible = True


            Case Else   '隱藏
                DAttachfile1.Visible = False
                DChkData1.Visible = False

        End Select

        Select Case FindFieldInf("Attachfile2")
            Case 0  '顯示            

                DAttachfile2.Visible = True
                DChkData2.Visible = True

            Case 1  '修改+檢查
                DAttachfile2.Visible = True
                DChkData2.Visible = True
            Case 2  '修改
                DAttachfile2.Visible = True
                DChkData2.Visible = True


            Case Else   '隱藏
                DAttachfile2.Visible = False
                DChkData2.Visible = False

        End Select


        Select Case FindFieldInf("Attachfile3")
            Case 0  '顯示            

                DAttachfile3.Visible = True
                DChkData3.Visible = True


                DAttachfile6.Visible = True
                DChkData6.Visible = True


            Case 1  '修改+檢查
                DAttachfile3.Visible = True
                DChkData3.Visible = True

                DAttachfile6.Visible = True
                DChkData6.Visible = True


            Case 2  '修改
                DAttachfile3.Visible = True
                DChkData3.Visible = True

                DAttachfile6.Visible = True
                DChkData6.Visible = True


            Case Else   '隱藏
                DAttachfile3.Visible = False
                DChkData3.Visible = False

                DAttachfile6.Visible = False
                DChkData6.Visible = False

        End Select



        Select Case FindFieldInf("Attachfile4")
            Case 0  '顯示            

                DAttachfile4.Visible = True
                DChkData4.Visible = True

            Case 1  '修改+檢查
                DAttachfile4.Visible = True
                DChkData4.Visible = True
            Case 2  '修改
                DAttachfile4.Visible = True
                DChkData4.Visible = True


            Case Else   '隱藏
                DAttachfile4.Visible = False
                DChkData4.Visible = False

        End Select



        Select Case FindFieldInf("Attachfile5")
            Case 0  '顯示            

                DAttachfile5.Visible = True
                DChkData5.Visible = True

            Case 1  '修改+檢查
                DAttachfile5.Visible = True
                DChkData5.Visible = True
            Case 2  '修改
                DAttachfile5.Visible = True
                DChkData5.Visible = True


            Case Else   '隱藏
                DAttachfile5.Visible = False
                DChkData5.Visible = False
        End Select


        '備註
        Select Case FindFieldInf("REMARK1")
            Case 0  '顯示              
                DREMARK1.BackColor = Color.LightGray
                DREMARK1.Visible = True
                DREMARK1.Attributes.Add("readonly", "true")

                DREPORTDATE.BackColor = Color.LightGray
                DREPORTDATE.Visible = True
                DREPORTDATE.Attributes.Add("readonly", "true")
                BREPORTDATE.Visible = False

                DQCACCDATE.BackColor = Color.LightGray
                DQCACCDATE.Visible = True
                DQCACCDATE.Attributes.Add("readonly", "true")
                BQCACCDATE.Visible = False

            Case 1  '修改+檢查
                DREMARK1.Visible = True
                DREMARK1.BackColor = Color.GreenYellow
                DREMARK1.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK1Rqd", "DREMARK1", "異常：需輸入備註")

                DREPORTDATE.Visible = True
                DREPORTDATE.BackColor = Color.GreenYellow
                DREPORTDATE.ReadOnly = False
                ShowRequiredFieldValidator("DReportDateRqd", "DReportDate", "異常：需輸入報表期限")
                BREPORTDATE.Visible = True

                DQCACCDATE.Visible = True
                DQCACCDATE.BackColor = Color.GreenYellow
                DQCACCDATE.ReadOnly = False
                ShowRequiredFieldValidator("DQCACCDATERqd", "DQCACCDATE", "異常：需輸入品管受理日")
                BQCACCDATE.Visible = True

            Case 2  '修改
                DREMARK1.Visible = True
                DREMARK1.BackColor = Color.Yellow
                DREMARK1.ReadOnly = False

                DREPORTDATE.Visible = True
                DREPORTDATE.BackColor = Color.Yellow
                DREPORTDATE.ReadOnly = False
                BREPORTDATE.Visible = True

                DQCACCDATE.Visible = True
                DQCACCDATE.BackColor = Color.Yellow
                DQCACCDATE.ReadOnly = False
                BQCACCDATE.Visible = True

            Case Else   '隱藏
                DREMARK1.Visible = False
                DREPORTDATE.Visible = False
                BREPORTDATE.Visible = False
                DQCACCDATE.Visible = False
                BQCACCDATE.Visible = False

        End Select


        '中文內容
        Select Case FindFieldInf("REMARK1")
            Case 0  '顯示              
                DQCCDESC.BackColor = Color.LightGray
                DQCCDESC.Visible = True
                DQCCDESC.Attributes.Add("readonly", "true")
                DQCCDESC.BackColor = Color.LightGray
                DQCCDESC.Visible = True
                DQCCDESC.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCCDESC.Visible = True
                DQCCDESC.BackColor = Color.GreenYellow
                DQCCDESC.ReadOnly = False
                ShowRequiredFieldValidator("DQCCDESCRqd", "DQCCDESC", "異常：需輸入中文內容")
            Case 2  '修改
                DQCCDESC.Visible = True
                DQCCDESC.BackColor = Color.Yellow
                DQCCDESC.ReadOnly = False
            Case Else   '隱藏
                DQCCDESC.Visible = False
        End Select


        '中文內容
        Select Case FindFieldInf("REMARK1")
            Case 0  '顯示              
                DQCEDESC.BackColor = Color.LightGray
                DQCEDESC.Visible = True
                DQCEDESC.Attributes.Add("readonly", "true")
                DQCEDESC.BackColor = Color.LightGray
                DQCEDESC.Visible = True
                DQCEDESC.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQCEDESC.Visible = True
                DQCEDESC.BackColor = Color.GreenYellow
                DQCEDESC.ReadOnly = False
                ShowRequiredFieldValidator("DQCEDESCRqd", "DQCEDESC", "異常：需輸入英文內容")
            Case 2  '修改
                DQCEDESC.Visible = True
                DQCEDESC.BackColor = Color.Yellow
                DQCEDESC.ReadOnly = False
            Case Else   '隱藏
                DQCEDESC.Visible = False
        End Select



        '分析依賴  20220316
        Select Case FindFieldInf("REMARK1")
            Case 0  '顯示           
                DMAT.Visible = True
                DRELY.Visible = True

            Case 1  '修改+檢查
                DMAT.Visible = True
                DRELY.Visible = True
            Case 2  '修改
                DMAT.Visible = True
                DRELY.Visible = True

            Case Else   '隱藏
                DMAT.Visible = False
                DRELY.Visible = False
        End Select


        '備註
        Select Case FindFieldInf("REMARK2")
            Case 0  '顯示              
                DREMARK2.BackColor = Color.LightGray
                DREMARK2.Visible = True
                DREMARK2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DREMARK2.Visible = True
                DREMARK2.BackColor = Color.GreenYellow
                DREMARK2.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK2Rqd", "DREMARK2", "異常：需輸入恆久對策")
            Case 2  '修改
                DREMARK2.Visible = True
                DREMARK2.BackColor = Color.Yellow
                DREMARK2.ReadOnly = False
            Case Else   '隱藏
                DREMARK2.Visible = False
        End Select

        '備註
        Select Case FindFieldInf("REMARK3")
            Case 0  '顯示              
                DREMARK3.BackColor = Color.LightGray
                DREMARK3.Visible = True
                DREMARK3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DREMARK3.Visible = True
                DREMARK3.BackColor = Color.GreenYellow
                DREMARK3.ReadOnly = False
                ShowRequiredFieldValidator("DREMARK3Rqd", "DREMARK3", "異常：需輸入未然防止3奌SET內容確認")
            Case 2  '修改
                DREMARK3.Visible = True
                DREMARK3.BackColor = Color.Yellow
                DREMARK3.ReadOnly = False
            Case Else   '隱藏
                DREMARK3.Visible = False
        End Select

        'Division2
        Select Case FindFieldInf("CPQTYCHK")
            Case 0  '顯示

                DCPQTYCHK.Enabled = False
            Case 1  '修改+檢查

                ShowRequiredFieldValidator("DCPQTYCHKRqd", "DCPQTYCHK", "異常：客訴數量確認中")

                DCPQTYCHK.Visible = True
            Case 2  '修改

                DCPQTYCHK.Visible = True
            Case Else   '隱藏

                DCPQTYCHK.Visible = False
        End Select



        '
        Select Case FindFieldInf("Attachfile1")
            Case 0  '顯示            

                DAttachfile1.Visible = True
                DChkData1.Visible = True

            Case 1  '修改+檢查
                DAttachfile1.Visible = True
                DChkData1.Visible = True
            Case 2  '修改
                DAttachfile1.Visible = True
                DChkData1.Visible = True


            Case Else   '隱藏
                DAttachfile1.Visible = False
                DChkData1.Visible = False

        End Select

        Select Case FindFieldInf("Attachfile2")
            Case 0  '顯示            

                DAttachfile2.Visible = True
                DChkData2.Visible = True

            Case 1  '修改+檢查
                DAttachfile2.Visible = True
                DChkData2.Visible = True
            Case 2  '修改
                DAttachfile2.Visible = True
                DChkData2.Visible = True


            Case Else   '隱藏
                DAttachfile2.Visible = False
                DChkData2.Visible = False

        End Select



    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("ComplaintOutPath")



        Dim SQL As String
        SQL = "Select * From F_ComplaintOutSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            '表單資料
            DNO.Text = dtData.Rows(0).Item("No")
            DDATE.Text = dtData.Rows(0).Item("DATE")
            SetFieldData("GLOBAL", dtData.Rows(0).Item("GLOBAL"))
            DCUSTOMERCODE.Text = dtData.Rows(0).Item("CUSTOMERCODE")
            DCUSTOMER.Text = dtData.Rows(0).Item("CUSTOMER")
            SetFieldData("APPNAME", dtData.Rows(0).Item("APPNAME"))

            DOORNO.Text = dtData.Rows(0).Item("OORNO")
            DNORNO.Text = dtData.Rows(0).Item("NORNO")
            DSPEC.Text = dtData.Rows(0).Item("SPEC")
            DSPECNAME.Text = dtData.Rows(0).Item("SPECNAME")
            DORQTY.Text = dtData.Rows(0).Item("ORQTY")


            If Mid(dtData.Rows(0).Item("ORDERDATE").ToString, 1, 4) = "1900" Then
                DORDERDATE.Text = ""
            Else
                DORDERDATE.Text = dtData.Rows(0).Item("ORDERDATE")
            End If

            If Mid(dtData.Rows(0).Item("SHIPDATE").ToString, 1, 4) = "1900" Then
                DSHIPDATE.Text = ""
            Else
                DSHIPDATE.Text = dtData.Rows(0).Item("SHIPDATE")
            End If
            SetFieldData("BIGGOODS", dtData.Rows(0).Item("BIGGOODS"))
            SetFieldData("PLACE", dtData.Rows(0).Item("PLACE"))
            If dtData.Rows(0).Item("ITEM") <> "" Then
                SetFieldData("ITEM", dtData.Rows(0).Item("ITEM"))
            End If

            DMFGDATE.Text = dtData.Rows(0).Item("MFGDATE")

            DCPCONTENT.Text = dtData.Rows(0).Item("CPCONTENT")
            DEPCONTENT.Text = dtData.Rows(0).Item("EPCONTENT")

            DCPQTY.Text = dtData.Rows(0).Item("CPQTY")
            If dtData.Rows(0).Item("CPQTYCHK") = 1 Then
                DCPQTYCHK.Checked = True
            Else
                DCPQTYCHK.Checked = False
            End If
            DBUYERCODE.Text = dtData.Rows(0).Item("BUYERCODE")
            DBUYER.Text = dtData.Rows(0).Item("BUYER")
            DMappath.Text = Path & dtData.Rows(0).Item("MapFile")

            '不良品樣品圖
            If Trim(dtData.Rows(0).Item("MapFile")) <> "" Then

                LMapFile.ImageUrl = DMappath.Text
                LMapFile1.NavigateUrl = DMappath.Text
                LMapFile.Visible = True
                LMapFile1.Visible = True
            Else
                LMapFile.Visible = False
                LMapFile1.Visible = False
            End If

            ' SetFieldData("ITEM", dtData.Rows(0).Item("ITEM"))
            SetFieldData("REPLAYLAN", dtData.Rows(0).Item("REPLAYLAN"))
            DQCNO.Text = dtData.Rows(0).Item("QCNO")

            SetFieldData("TYPE", dtData.Rows(0).Item("TYPE"))

            SetFieldData("PL", dtData.Rows(0).Item("PL"))
            SetFieldData("ACCDEP1", dtData.Rows(0).Item("ACCDEP1"))

            If dtData.Rows(0).Item("ACCDEP1") <> "" Then


                If dtData.Rows(0).Item("ACCDEP12") <> "" Then
                    SetFieldData("ACCDEP12", dtData.Rows(0).Item("ACCDEP12"))
                End If
                If dtData.Rows(0).Item("ACCDEP13") <> "" Then
                    SetFieldData("ACCDEP13", dtData.Rows(0).Item("ACCDEP13"))
                End If

            End If


            SetFieldData("TYPE1", dtData.Rows(0).Item("TYPE1"))
            SetFieldData("TYPE2", dtData.Rows(0).Item("TYPE2"))
            SetFieldData("QCDESCTYPE", dtData.Rows(0).Item("QCDESCTYPE"))
      


            DACCEMPNAME.Text = dtData.Rows(0).Item("ACCEMPNAME")

            SetFieldData("ACCDEP2", dtData.Rows(0).Item("ACCDEP2"))
            DACCDEP2NO.Text = dtData.Rows(0).Item("ACCDEP2NO")

            SetFieldData("CUSTOMERTYPE", dtData.Rows(0).Item("CUSTOMERTYPE"))
            SetFieldData("JUDGE", dtData.Rows(0).Item("JUDGE"))
            SetFieldData("MANUREASON", dtData.Rows(0).Item("MANUREASON"))
            SetFieldData("MANUREASON1", dtData.Rows(0).Item("MANUREASON1"))

            SetFieldData("RESPONS", dtData.Rows(0).Item("RESPONS"))

            DHOUR1.Text = dtData.Rows(0).Item("HOUR1")
            DHOUR2.Text = dtData.Rows(0).Item("HOUR2")
            DHOUR3.Text = dtData.Rows(0).Item("HOUR3")
            DHOUR4.Text = dtData.Rows(0).Item("HOUR4")
            DHOUR5.Text = dtData.Rows(0).Item("HOUR5")

            DREMARK1.Text = dtData.Rows(0).Item("REMARK1")
            DREMARK2.Text = dtData.Rows(0).Item("REMARK2")
            DREMARK3.Text = dtData.Rows(0).Item("REMARK3")

            DQCCDESC.Text = dtData.Rows(0).Item("QCCDESC")
            DQCEDESC.Text = dtData.Rows(0).Item("QCEDESC")

            If dtData.Rows(0).Item("MACH") = "0" Then
                DMach1.Checked = True
                DMach2.Checked = False
            ElseIf dtData.Rows(0).Item("MACH") = "1" Then
                DMach2.Checked = True
                DMach1.Checked = False
            Else
                DMach1.Checked = False
                DMach2.Checked = False
            End If

            DMACHNO.Text = dtData.Rows(0).Item("MACHNO")


            SetFieldData("HAPPEN", dtData.Rows(0).Item("HAPPEN"))

            'jessica 20220323 轉分析依賴
            If dtData.Rows(0).Item("RELY") = 1 Then
                DRELY.Checked = True
            Else
                DRELY.Checked = False
            End If

            'jessica 20230213 物料抱怨單
            If dtData.Rows(0).Item("MAT") = 1 Then
                DMAT.Checked = True
            Else
                DMAT.Checked = False
            End If



            If Mid(dtData.Rows(0).Item("SAMPLEDATE").ToString, 1, 4) = "1900" Then
                DSAMPLEDATE.Text = ""
            Else
                DSAMPLEDATE.Text = dtData.Rows(0).Item("SAMPLEDATE")
            End If


            If Mid(dtData.Rows(0).Item("QCDATE").ToString, 1, 4) = "1900" Then
                DQCDATE.Text = ""
            Else
                DQCDATE.Text = dtData.Rows(0).Item("QCDATE")
            End If

            DANSWERDAYS.Text = dtData.Rows(0).Item("ANSWERDAYS")
            SetFieldData("REMARK", dtData.Rows(0).Item("REMARK"))

            DACOST.Text = dtData.Rows(0).Item("ACOST")
            DBCOST.Text = dtData.Rows(0).Item("BCOST")

            DCCOST.Text = dtData.Rows(0).Item("CCOST")
            DDCOST.Text = dtData.Rows(0).Item("DCOST")

            DECOST.Text = dtData.Rows(0).Item("ECOST")
            DFCOST.Text = dtData.Rows(0).Item("FCOST")
            DGCOST.Text = dtData.Rows(0).Item("GCOST")
            DHCOST.Text = dtData.Rows(0).Item("HCOST")

            DALLCOST.Text = dtData.Rows(0).Item("ALLCOST")

            SetFieldData("POINT1", dtData.Rows(0).Item("POINT1"))
            DPOINT2.Text = dtData.Rows(0).Item("POINT2")
            DPOINT3.Text = dtData.Rows(0).Item("POINT3")


            SetFieldData("SCORE1", dtData.Rows(0).Item("SCORE1"))
            SetFieldData("SCORE2", dtData.Rows(0).Item("SCORE2"))
            SetFieldData("SCORE3", dtData.Rows(0).Item("SCORE3"))


            DCOST.Text = dtData.Rows(0).Item("COST")


            If Mid(dtData.Rows(0).Item("CUSTDATE").ToString, 1, 4) = "1900" Then
                DCUSTDATE.Text = ""
            Else
                DCUSTDATE.Text = dtData.Rows(0).Item("CUSTDATE")
            End If

            If Mid(dtData.Rows(0).Item("YFGCDATE").ToString, 1, 4) = "1900" Then
                DYFGCDATE.Text = ""
            Else
                DYFGCDATE.Text = dtData.Rows(0).Item("YFGCDATE")
            End If

            '人工薪
            DWORKPAY.Text = dtData.Rows(0).Item("workpay")


            'sql = sql + " DELIVERYDATE,SAMPLE,TYPE1,TYPE2,ReportDate,QCDESCTYPE,JUDGE1,JUDGE2,CARDATE,FIRSTDATE,MIDDATE,"

            If Mid(dtData.Rows(0).Item("DELIVERYDATE").ToString, 1, 4) = "1900" Then
                DDELIVERYDATE.Text = ""
            Else
                DDELIVERYDATE.Text = dtData.Rows(0).Item("DELIVERYDATE")

            End If

            SetFieldData("SAMPLE", dtData.Rows(0).Item("SAMPLE"))


            SetFieldData("QCDESCTYPE,", dtData.Rows(0).Item("QCDESCTYPE"))

            If dtData.Rows(0).Item("JUDGE") <> "" Then
                If dtData.Rows(0).Item("JUDGE1") <> "" Then
                    SetFieldData("JUDGE1", dtData.Rows(0).Item("JUDGE1"))
                End If


                If dtData.Rows(0).Item("JUDGE1") <> "" Then
                    SetFieldData("JUDGE2", dtData.Rows(0).Item("JUDGE2"))

                End If

            End If

            If Mid(dtData.Rows(0).Item("REPORTDATE").ToString, 1, 4) = "1900" Then
                DREPORTDATE.Text = ""
            Else
                DREPORTDATE.Text = dtData.Rows(0).Item("REPORTDATE")
            End If


            If Mid(dtData.Rows(0).Item("CARDATE").ToString, 1, 4) = "1900" Then
                DCARDATE.Text = ""
            Else
                DCARDATE.Text = dtData.Rows(0).Item("CARDATE")
            End If

            If Mid(dtData.Rows(0).Item("FIRSTDATE").ToString, 1, 4) = "1900" Then
                DFIRSTDATE.Text = ""
            Else
                DFIRSTDATE.Text = dtData.Rows(0).Item("FIRSTDATE")
            End If

            If Mid(dtData.Rows(0).Item("FINALDATE").ToString, 1, 4) = "1900" Then
                DFIRSTDATE.Text = ""
            Else
                DFINALDATE.Text = dtData.Rows(0).Item("FINALDATE")
            End If

            If Mid(dtData.Rows(0).Item("MIDDATE").ToString, 1, 4) = "1900" Then
                DMIDDATE.Text = ""
            Else
                DMIDDATE.Text = dtData.Rows(0).Item("MIDDATE")
            End If

            If Mid(dtData.Rows(0).Item("QCACCDATE").ToString, 1, 4) = "1900" Then
                DQCACCDATE.Text = ""
            Else
                DQCACCDATE.Text = dtData.Rows(0).Item("QCACCDATE")
            End If










            '判斷重要點數
            Total()

            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3109'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNO.Text + "/營業"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "\\" + DBAdapter3.Rows(0).Item("Data1") + DNO.Text + "\營業"

            Dim dirInfo As New System.IO.DirectoryInfo(OpenDir2)
            chktemp.Text = OpenDir2

            Dim FileDir As Integer  '資料夾
            FileDir = dirInfo.GetDirectories("*").Length
            Dim FileCount As Integer '檔案
            FileCount = dirInfo.GetFiles("*.*").Length


            If FileCount > 0 Or FileDir > 0 Then
                DChkData1.Checked = True

            Else
                DChkData1.Checked = False

            End If


            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")
            '  DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir2 + "','_blank');return false;")


            '交易資料
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dtWaitHandle.Rows.Count > 0 Then
                DDecideDesc.Text = dtWaitHandle.Rows(0)("DecideDesc")           '說明

                If dtWaitHandle.Rows(0)("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", dtWaitHandle.Rows(0)("ReasonCode"))    '延遲理由代碼
                        If dtWaitHandle.Rows(0)("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = dtWaitHandle.Rows(0)("Reason")  '延遲理由
                            DReasonDesc.Text = dtWaitHandle.Rows(0)("ReasonDesc")     '延遲其他說明
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
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            GridView2.DataSource = uDataBase.GetDataTable(SQL)
            GridView2.DataBind()
        End If
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        Dim idx As Integer = FindFieldInf(pFieldName)
        Dim i As Integer
        '擔當者及部門 

        sql = "Select Divname,Username From M_Users "
        sql = sql & " Where UserID = '" & wApplyID & "'"
        sql = sql & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(sql)


        '區域
        If pFieldName = "GLOBAL" Then
            DGLOBAL.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DGLOBAL.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'GLOBAL'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DGLOBAL.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DGLOBAL.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        '營業人員
        If pFieldName = "APPNAME" Then
            DAPPNAME.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAPPNAME.Items.Add(ListItem1)
                End If
            Else
                'sql = "  select substring(data,s+1,e-s-1)Data from ("
                'sql = sql & " Select DATA,CHARINDEX('-',DATA)s,CHARINDEX('#',DATA)e from M_referp"
                sql = " select data from M_referp "
                sql = sql & " where  cat = '3109'"
                sql = sql & "and dkey = 'APPNAME'"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DAPPNAME.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAPPNAME.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'SAMPLE
        If pFieldName = "SAMPLE" Then
            DSAMPLE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSAMPLE.Items.Add(ListItem1)
                End If
            Else
                'sql = "  select substring(data,s+1,e-s-1)Data from ("
                'sql = sql & " Select DATA,CHARINDEX('-',DATA)s,CHARINDEX('#',DATA)e from M_referp"
                sql = " select data from M_referp "
                sql = sql & " where  cat = '3109'"
                sql = sql & "and dkey = 'SAMPLE'"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSAMPLE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSAMPLE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'BIGGOODS
        If pFieldName = "BIGGOODS" Then
            DBIGGOODS.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBIGGOODS.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'BIGGOODS'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DBIGGOODS.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBIGGOODS.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '縫製地點
        If pFieldName = "PLACE" Then
            DPLACE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPLACE.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'PLACE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DPLACE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPLACE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'CUSTOMERTYPE
        If pFieldName = "CUSTOMERTYPE" Then
            DCUSTOMERTYPE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCUSTOMERTYPE.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'CUSTOMERTYPE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DCUSTOMERTYPE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCUSTOMERTYPE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'ITEM
        If pFieldName = "ITEM" Then
            DITEM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DITEM.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'ITEM'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DITEM.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DITEM.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'REPLAYLAN
        If pFieldName = "REPLAYLAN" Then
            DREPLAYLAN.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DREPLAYLAN.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'REPLAYLAN'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DREPLAYLAN.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DREPLAYLAN.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'ACCDEP1
        If pFieldName = "ACCDEP1" Then
            DACCDEP1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DACCDEP1.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  distinct substring(data,1,CHARINDEX('-',data)-1)Data  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'RDIVISION'  "

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DACCDEP1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DACCDEP1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'ACCDEP12
        If pFieldName = "ACCDEP12" Then
            DACCDEP12.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DACCDEP12.Items.Add(ListItem1)
                End If
            Else
                sql = "  select data,substring(dkey,8,len(dkey)-1) from  M_referp"
                sql = sql & " where  left(Dkey,6)='WhereP'"
                sql = sql & " and substring(dkey,8,len(dkey)-1) in("
                sql = sql & " select  substring(dkey,10,len(dkey)-1)Data from  M_referp where cat ='3109' and left(Dkey,8) ='WhereDep'"
                sql = sql & " and data like '%" + Trim(DACCDEP1.SelectedValue) + "%') order by Unique_id"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DACCDEP12.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DACCDEP12.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'ACCDEP13
        If pFieldName = "ACCDEP13" Then
            DACCDEP13.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DACCDEP13.Items.Add(ListItem1)
                End If
            Else
                sql = "  select  distinct data  from  M_referp where cat ='3109' and left(Dkey,4) ='What'"
                sql = sql & " and   left(Dkey,4)='What'"
                sql = sql & " and substring(dkey,6,len(dkey)-1) ='" + Trim(DACCDEP12.SelectedValue) + "' order by data "

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DACCDEP13.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DACCDEP13.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        'TYPE
        If pFieldName = "TYPE" Then
            DTYPE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTYPE.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select data  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'TYPE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DTYPE.Items.Add("")
                'DTYPE.Items.Add("依賴")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTYPE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'TYPE1
        If pFieldName = "TYPE1" Then
            DTYPE1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTYPE1.Items.Add(ListItem1)
                End If
            Else



                If DTYPE.SelectedValue <> "依賴" Then
                    sql = "  Select DATA from M_referp"
                    sql = sql & " where  cat = '3109'"
                    sql = sql & " and dkey = 'JUDGE'"
                    sql = sql & " order by unique_id"
                Else
                    sql = " SELECT*  from M_referp where cat ='3109' and DKEY ='RELY' "
                End If


                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DTYPE1.Items.Add("")
                'DTYPE1.Items.Add("依賴")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTYPE1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'TYPE2
        If pFieldName = "TYPE2" Then
            DTYPE2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTYPE2.Items.Add(ListItem1)
                End If
            Else




                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'RDIVISION1'"
                sql = sql & " order by unique_id"



                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DTYPE2.Items.Add("")
                'DTYPE2.Items.Add("依賴")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTYPE2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
            End If





        'QCDESCTYPE
        If pFieldName = "QCDESCTYPE" Then
            DQCDESCTYPE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCDESCTYPE.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select data  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'QCDESCTYPE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DQCDESCTYPE.Items.Add("")
                'DQCDESCTYPE.Items.Add("依賴")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCDESCTYPE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'PL
        If pFieldName = "PL" Then
            DPL.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPL.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select data  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'PL'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DPL.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPL.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If







        'ACCDEP2
        If pFieldName = "ACCDEP2" Then
            DACCDEP2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DACCDEP2.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select substring(data,9,len(data)-1) data1,substring(data,1,7) data2  from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and left(dkey,3) = 'DEP'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DACCDEP2.Items.Add("")

                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data1")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data1")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DACCDEP2.Items.Add(ListItem1)
                Next
                'If dtReferp.Rows.Count = 0 Then
                '    DACCDEP2.BackColor = Color.Yellow
                '    'ShowRequiredFieldValidator("DACCDEP2Rqd", "DACCDEP2", "異常：需輸入責任工程別")
                '    DACCDEP2.Visible = False
                'Else
                '    DACCDEP2.BackColor = Color.GreenYellow
                '    'ShowRequiredFieldValidator("DACCDEP2Rqd", "DACCDEP2", "異常：需輸入責任工程別")
                '    DACCDEP2.Visible = True
                'End If
                dtReferp.Clear()

            End If
        End If




        'JUDGE
        If pFieldName = "JUDGE" Then
            DJUDGE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DJUDGE.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'JUDGE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DJUDGE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DJUDGE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'JUDGE1
        If pFieldName = "JUDGE1" Then
            DJUDGE1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DJUDGE1.Items.Add(ListItem1)
                End If
            Else
                sql = "  SELECT data  FROM M_referp where cat ='3109' and left(dkey,9)='JUDGETYPE' and substring(dkey,11,1) ='" + Mid(DJUDGE.SelectedValue, 1, 1) + "' order by Unique_id "


                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DJUDGE1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DJUDGE1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'JUDGE2
        If pFieldName = "JUDGE2" Then
            DJUDGE2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DJUDGE2.Items.Add(ListItem1)
                End If
            Else
              
                sql = "  SELECT  Data  FROM M_referp where cat ='3109' and  substring(dkey,9,len(dkey)-1) ='" + Trim(DJUDGE1.SelectedValue) + "' order by Unique_id "


                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DJUDGE2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DJUDGE2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If






        'MANUREASON
        If pFieldName = "MANUREASON" Then
            DMANUREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMANUREASON.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'MANUREASON'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DMANUREASON.Items.Clear()
                DMANUREASON.Items.Add("")

                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMANUREASON.Items.Add(ListItem1)
                Next

                dtReferp.Clear()


            End If
        End If


        'MANUREASON
        If pFieldName = "MANUREASON1" Then
            DMANUREASON1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMANUREASON1.Items.Add(ListItem1)
                End If
            Else
                If DMANUREASON.SelectedValue = "1.人為" Then
                    sql = "  Select DATA from M_referp"
                    sql = sql & " where  cat = '3109'"
                    sql = sql & " and dkey = 'MANUREASON1'"
                    sql = sql & " order by unique_id"

                    Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                    DMANUREASON1.Items.Clear()
                    DMANUREASON1.Items.Add("")
                    For i = 0 To dtReferp.Rows.Count - 1
                        Dim ListItem1 As New ListItem
                        ListItem1.Text = dtReferp.Rows(i).Item("Data")
                        ListItem1.Value = dtReferp.Rows(i).Item("Data")
                        If ListItem1.Value = pName Then ListItem1.Selected = True
                        DMANUREASON1.Items.Add(ListItem1)
                    Next
                    dtReferp.Clear()
                Else
                    DMANUREASON1.Items.Clear()
                    DMANUREASON1.Items.Add("無")
                End If


            End If
        End If




        'RESPONS
        If pFieldName = "RESPONS" Then
            DRESPONS.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DRESPONS.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'RESPONS'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DRESPONS.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DRESPONS.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'REMARK
        If pFieldName = "REMARK" Then
            DREMARK.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DREMARK.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'REMARK'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DREMARK.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DREMARK.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'POINT1
        If pFieldName = "POINT1" Then
            DPOINT1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPOINT1.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'POINT1'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DPOINT1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPOINT1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'SCORE1
        If pFieldName = "SCORE1" Then
            DSCORE1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSCORE1.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'SCORE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSCORE1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSCORE1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'SCORE2
        If pFieldName = "SCORE2" Then
            DSCORE2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSCORE2.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'SCORE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSCORE2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSCORE2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'SCORE3
        If pFieldName = "SCORE3" Then
            DSCORE3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSCORE3.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'SCORE'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSCORE3.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSCORE3.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'Happen
        If pFieldName = "HAPPEN" Then
            DHappen.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DHappen.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select DATA from M_referp"
                sql = sql & " where  cat = '3109'"
                sql = sql & " and dkey = 'Happen'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DHappen.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DHappen.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
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
                sql = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReasonCode.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReasonCode.Rows(i)("DKey")
                    ListItem1.Value = dtReasonCode.Rows(i)("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由
        If pFieldName = "Reason" Then
            sql = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim dtReason As DataTable = uDataBase.GetDataTable(sql)
            For i = 0 To dtReason.Rows.Count - 1
                DReason.Text = dtReason.Rows(i)("Data")
            Next
        End If
    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        Top = 2050
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
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
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()

        BCustomer.Attributes.Add("onClick", "GetCustomer()") '找客戶
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '找buyer
        BSPEC.Attributes.Add("onClick", "GetSPEC()") '找buyer

        BDATE.Attributes("onclick") = "calendarPicker('Form1.DDATE');"
        BORDERDATE.Attributes("onclick") = "calendarPicker('Form1.DORDERDATE');"
        BSHIPDATE.Attributes("onclick") = "calendarPicker('Form1.DSHIPDATE');"
        BSAMPLEDATE.Attributes("onclick") = "calendarPicker('Form1.DSAMPLEDATE');"
        BQCDATE.Attributes("onclick") = "calendarPicker('Form1.DQCDATE');"
        BCUSTDATE.Attributes("onclick") = "calendarPicker('Form1.DCUSTDATE');"
        BYFGCDATE.Attributes("onclick") = "calendarPicker('Form1.DYFGCDATE');"
        BREPORTDATE.Attributes("onclick") = "calendarPicker('Form1.DREPORTDATE');"
        BDELIVERYDATE.Attributes("onclick") = "calendarPicker('Form1.DDELIVERYDATE');"
        BQCACCDATE.Attributes("onclick") = "calendarPicker('Form1.DQCACCDATE');"
        BREPORTDATE.Attributes("onclick") = "calendarPicker('Form1.DREPORTDATE');"
        BCARDATE.Attributes("onclick") = "calendarPicker('Form1.DCARDATE');"
        BFIRSTDATE.Attributes("onclick") = "calendarPicker('Form1.DFIRSTDATE');"
        BFINALDATE.Attributes("onclick") = "calendarPicker('Form1.DFINALDATE');"
        BMIDDATE.Attributes("onclick") = "calendarPicker('Form1.DMIDDATE');"
        BMFGDate.Attributes("onclick") = "calendarPicker('Form1.DMFGDATE');"
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
        Dim Run As Boolean = True           '是否執行
        Dim RepeatRun As Boolean = False    '是否重覆執行
        Dim MultiJob As Integer = 0         '多人核定
        Dim wLevel As String = ""           '難易度
        '
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
                        '申請流程資料建置(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                        oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '設定委託No
                        DNO.Text = SetNo(NewFormSno)
                    End If
                    pRunNextStep = 1
                Else
                    If RepeatRun = False Then   '不是通知的重覆執行
                        '更新交易資料
                        ModifyTranData(pFun, pSts)
                        '流程資料結束(表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽))
                        RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)

                        'jessica 2021/01/04 強迫變1
                        If wStep = 10 And pFun = "NG1" Then
                            pRunNextStep = 1
                        End If

                        If RtnCode <> 0 Then
                            ErrCode = 9120
                        End If
                    Else
                        pRunNextStep = 1    '是通知的重覆執行
                    End If
                End If
            End If


            '取得下一關
            If ErrCode = 0 And pRunNextStep = 1 Then
                '指定簽核者User ID
                Dim wAllocateID As String = ""

                '   If wStep = 10 And pFun = "OK" Then  '品管送給製造擔當
                'If DTYPE.SelectedValue = "一般" Then
                'pAction = 0
                'wAllocateID = GetUerID(Trim(DACCEMPNAME.Text))  '被訴擔當
                'Else
                '   pAction = 2
                'End If
                'End If
                'If wStep = 15 Then
                '    wAllocateID = GetUerID(Trim(DACCEMPNAME.Text))  '被訴擔當
                'End If

                '製告擔當
                If wStep = 10 Then
                    If pFun = "OK" Then
                        If DTYPE.SelectedValue = "一般" Or DTYPE.SelectedValue = "一般特別監視" Then
                            pAction = 0  '11
                        ElseIf DTYPE.SelectedValue = "重大" Then
                            pAction = 2  '15
                        Else  '依賴
                            pAction = 3  '22
                        End If
                    Else 'ng
                        pAction = 1
                    End If
                End If



                '取得簽核者
                '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
                RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                              pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)


                ''申請後給品管判定
                'If wStep = 1 Or (wStep = 500 And pFun = "OK") Then

                '    Dim SQL1 As String
                '    SQL1 = " SELECT * FROM m_referp"
                '    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION1'"

                '    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '    pCount = dtReferp.Rows.Count
                '    For i = 0 To dtReferp.Rows.Count - 1
                '        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
                '    Next
                '    dtReferp.Clear()
                '    RtnCode = 0
                '    pRunNextStep = 1
                'End If


                ''品管確認後給製造部台日籍
                'If wStep = 10 And pFun = "OK" Then
                '    If DTYPE.SelectedValue = "一般" Then
                '        '判斷製造部門台日籍
                '        Dim SQL1 As String
                '        SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                '        SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION'"
                '        SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DACCDEP1.SelectedValue + "'"

                '        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '        pCount = dtReferp.Rows.Count
                '        For i = 0 To dtReferp.Rows.Count - 1
                '            pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                '        Next
                '        dtReferp.Clear()
                '        RtnCode = 0
                '        pRunNextStep = 1
                '    End If

                'End If


                ''廠長確認後給品管最後確認
                'If wStep = 25 And pFun = "OK" Then

                '    Dim SQL1 As String
                '    SQL1 = " SELECT * FROM m_referp"
                '    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION1'"

                '    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '    pCount = dtReferp.Rows.Count
                '    For i = 0 To dtReferp.Rows.Count - 1
                '        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
                '    Next
                '    dtReferp.Clear()
                '    RtnCode = 0
                '    pRunNextStep = 1
                'End If


                '第一階段不分營業課別及責任部門()

                '申請後給品管判定
                '品管判定/品管最後確認
                'If wStep = 18 Then
                '    Dim a As Integer
                '    a = 0

                'End If

                If wStep = 6 Or wStep = 26 Or wStep = 19 Or wStep = 22 Then
                    Dim SQL1 As String
                    SQL1 = " SELECT * FROM m_referp"
                    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION1'"

                    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                    pCount = dtReferp.Rows.Count
                    For i = 0 To dtReferp.Rows.Count - 1
                        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
                    Next
                    dtReferp.Clear()
                    RtnCode = 0
                    pRunNextStep = 1
                End If

                If (wStep = 20 Or wStep = 22) And pFun = "NG1" Then
                    Dim SQL1 As String
                    SQL1 = " SELECT * FROM m_referp"
                    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION1'"

                    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                    pCount = dtReferp.Rows.Count
                    For i = 0 To dtReferp.Rows.Count - 1
                        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA"))
                    Next
                    dtReferp.Clear()
                    RtnCode = 0
                    pRunNextStep = 1
                End If

                '品管確認後給製造部台日籍
                If wStep = 12 Or wStep = 15 Then
                    ' If DTYPE.SelectedValue = "一般" Then
                    '判斷製造部門台日籍
                    Dim SQL1 As String
                    SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION'"
                    SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DACCDEP1.SelectedValue + "'"

                    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                    pCount = dtReferp.Rows.Count
                    For i = 0 To dtReferp.Rows.Count - 1
                        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                    Next
                    dtReferp.Clear()
                    RtnCode = 0
                    pRunNextStep = 1
                    'End If

                End If

                '廠長NG退回給製造 20240710
                If wStep = 25 And pFun = "NG1" Then
                    ' If DTYPE.SelectedValue = "一般" Then
                    '判斷製造部門台日籍
                    Dim SQL1 As String
                    SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION'"
                    SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DACCDEP1.SelectedValue + "'"

                    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                    pCount = dtReferp.Rows.Count
                    For i = 0 To dtReferp.Rows.Count - 1
                        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                    Next
                    dtReferp.Clear()
                    RtnCode = 0
                    pRunNextStep = 1
                    'End If
                End If



                Dim AccdepSQL As String = ""
                If DTYPE.SelectedValue = "依賴" Then

                    If DACCDEP1.SelectedValue <> "" Then
                        AccdepSQL = "'" + DACCDEP1.SelectedValue + "'"
                    End If
                    'If DACCDEP12.SelectedValue <> "" Then
                    '    AccdepSQL = AccdepSQL + ",'" + DACCDEP12.SelectedValue + "'"
                    'End If
                    'If DACCDEP13.SelectedValue <> "" Then
                    '    AccdepSQL = AccdepSQL + ",'" + DACCDEP13.SelectedValue + "'"
                    'End If

                End If

                ' 品管確認依賴後cc給製造部台日籍
                If wStep = 18 Or wStep = 38 Then
                    If DTYPE.SelectedValue = "依賴" Then
                        '判斷製造部門台日籍
                        Dim SQL1 As String
                        SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                        SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION'"
                        SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1) IN (" + AccdepSQL + ")"

                        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                        pCount = dtReferp.Rows.Count
                        For i = 0 To dtReferp.Rows.Count - 1
                            pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                        Next
                        dtReferp.Clear()
                        RtnCode = 0
                        pRunNextStep = 1
                    End If

                End If



                ''品管確認後給製造部台日籍
                'If  wStep = 14 or step =33  Then
                '    If DTYPE.SelectedValue = "一般" Then
                '        '判斷製造部門台日籍
                '        Dim SQL1 As String
                '        SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                '        SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION'"
                '        SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DACCDEP1.SelectedValue + "'"

                '        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '        pCount = dtReferp.Rows.Count
                '        For i = 0 To dtReferp.Rows.Count - 1
                '            pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                '        Next
                '        dtReferp.Clear()
                '        RtnCode = 0
                '        pRunNextStep = 1
                '    End If

                'End If



                '第二階段分營業課別及責任部門


                ''品管確認CC -營業台日籍主管
                'If wStep = 12 or wstep=32 Then
                '    Dim SQL1 As String
                '    SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                '    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION3'"
                '    SQL1 = SQL1 + " and SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DAPPNAME.SelectedValue + "'"

                '    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '    pCount = dtReferp.Rows.Count
                '    For i = 0 To dtReferp.Rows.Count - 1
                '        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                '    Next
                '    dtReferp.Clear()
                '    RtnCode = 0
                '    pRunNextStep = 1
                'End If


                ''品管確認CC -責任部門編輯人員工程:14
                'If wStep = 13 Then
                '    Dim SQL1 As String
                '    SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                '    SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION2'"
                '    SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DACCDEP1.SelectedValue + "'"

                '    Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '    pCount = dtReferp.Rows.Count
                '    For i = 0 To dtReferp.Rows.Count - 1
                '        pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                '    Next
                '    dtReferp.Clear()
                '    RtnCode = 0
                '    pRunNextStep = 1
                'End If


                ''品管確認後給製造部台日籍
                'If  wStep = 14 or step =33  Then
                '    If DTYPE.SelectedValue = "一般" Then
                '        '判斷製造部門台日籍
                '        Dim SQL1 As String
                '        SQL1 = " SELECT *,SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)DATA1,SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA2 FROM m_referp"
                '        SQL1 = SQL1 + "  WHERE DKEY = 'RDIVISION'"
                '        SQL1 = SQL1 + " AND SUBSTRING(DATA,1,CHARINDEX('-',DATA)-1)='" + DACCDEP1.SelectedValue + "'"

                '        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)
                '        pCount = dtReferp.Rows.Count
                '        For i = 0 To dtReferp.Rows.Count - 1
                '            pNextGate(i + 1) = GetUerID(dtReferp.Rows(i).Item("DATA2"))
                '        Next
                '        dtReferp.Clear()
                '        RtnCode = 0
                '        pRunNextStep = 1
                '    End If

                'End If






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
                        '取得核定者-群組行是曆
                        wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                        '取得工程負荷最後日時(核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數)
                        oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)

                        '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                        oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                        '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                        RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)

                        If RtnCode <> 0 Then
                            ErrCode = 9150
                            Exit For
                        End If
                    Next i
                Else
                    '流程結束(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                    RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)

                    If RtnCode <> 0 Then ErrCode = 9160
                End If
            End If
            '當工程日程調整
            If ErrCode = 0 Then
                If RepeatRun = False Then
                    '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                    RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
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
                        If wStep <> 6 Then
                            ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                        End If

                    End If
                    AddCommissionNo(wFormNo, wFormSno)
                End If
                '傳送郵件
                If pNextStep <> 999 Then
                    For i = 1 To pCount
                        ' If pNextStep <> 25 Then  '廠長工程不要寄信
                        oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        'End If

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
                uJavaScript.PopMsg(Me, Message)
            End If      '儲存表單ErrCode=0
        End While       '重覆執行

        If ErrCode = 0 Then
            '--郵件傳送---------
            oCommon.SendMail()
            '
            Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
            Response.Redirect(URL)

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)

        Dim SQL1 As String
        Dim WORKPAY As String = ""
        Dim i As Integer
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                                  CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                  System.Configuration.ConfigurationManager.AppSettings("ComplaintOutPath")


        SQL1 = " SELECT * FROM    m_referp"
        SQL1 = SQL1 + "  where cat = '3109'"
        SQL1 = SQL1 + " and dkey = 'WORKPAY'"
        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL1)

        For i = 0 To dtReferp.Rows.Count - 1
            WORKPAY = dtReferp.Rows(i).Item("Data")
        Next
        dtReferp.Clear()

        Dim FileName As String
        Dim sql As String = ""
        sql = " Insert into F_ComplaintOutSheet (Sts, CompletedTime, FormNo, FormSno,"
        sql = sql + " NO,DATE,GLOBAL,CUSTOMERCODE,CUSTOMER,APPNAME, "  '11
        sql = sql + " OORNO,NORNO,SPEC,SPECNAME,ORQTY,ORDERDATE,SHIPDATE," '11
        sql = sql + " BIGGOODS,PLACE,CPCONTENT,EPCONTENT,CPQTY,CPQTYCHK,BuyerCode,Buyer,MapFile," '10
        sql = sql + " ITEM,REPLAYLAN," '11
        sql = sql + " QCRDATE,QCNO,ACCDEP1,ACCDEP12,ACCDEP13,TYPE,PL,ACCEMPNAME," '14
        sql = sql + " ACCDEP2,ACCDEP2NO,CUSTOMERTYPE,JUDGE,MANURDATE,MANUREASON,MANUREASON1,RESPONS,HOUR1,HOUR2,HOUR3,HOUR4,HOUR5, "  '9
        sql = sql + " SAMPLEDATE,QCDATE,ANSWERDAYS,REMARK,ACOST,BCOST,CCOST,DCOST,ECOST,ALLCOST,"  '13
        sql = sql + " POINT1,POINT2,POINT3,COST,CUSTDATE,YFGCDATE,WORKPAY,QCSts,Remark1,Remark2,Remark3,SCORE1,SCORE2,SCORE3,RELY,MAT,QCCDESC,QCEDESC," '4
        sql = sql + " DELIVERYDATE,SAMPLE,TYPE1,TYPE2,ReportDate,QCDESCTYPE,JUDGE1,JUDGE2,CARDATE,FIRSTDATE,FINALDATE,MIDDATE,QCACCDATE,MFGDATE,"
        sql = sql + " CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + " VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日
        sql = sql + " '003109', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號

        'NO,DATE,GLOBAL,CUSTOMERCODE,CUSTOMER,APPNAME,ADDRESS,PONO,YKKGROUP,MATERIAL,TEL,
        sql = sql + " N'" + YKK.ReplaceString(DNO.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDATE.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DGLOBAL.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCUSTOMERCODE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCUSTOMER.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DAPPNAME.SelectedValue) + "', "


        'OORNO,NORNO,REBACKNO,SPEC,ORDERDATE,SHIPDATE,ORQTY,ORQTY1,ORPRICE,SAMPLE,SAMPLEBDATE,
        sql = sql + " N'" + YKK.ReplaceString(DOORNO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DNORNO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSPEC.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSPECNAME.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DORQTY.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DORDERDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSHIPDATE.Text) + "', "



        'BIGGOODS,PLACE,WASHING,WASHCONDITION,CPCONTENT,CPQTY,CHECKQTY,CPMONEY,BuyerCode,Buyer,
        sql = sql + " N'" + YKK.ReplaceString(DBIGGOODS.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DPLACE.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCPCONTENT.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DEPCONTENT.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCPQTY.Text) + "', "
        If DCPQTYCHK.Checked Then
            sql = sql + " 1, "
        Else
            sql = sql + " 0, "
        End If
        sql = sql + " N'" + YKK.ReplaceString(DBUYERCODE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DBUYER.Text) + "', "

        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                'FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & System.IO.Path.GetExtension(DMapFile.PostedFile.FileName)

                DMapFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        sql = sql + " '" + FileName + "'," '不良樣品圖

        'CLAIM,ITEM,DELIVERYDATE,ADELIVERYDATE,ASKCONTENT,CHANGEQTY,CHANGENO,REPLAY,REPLAYLAN1,REPLAYLAN2,REPLAYLAN3

        sql = sql + " N'" + YKK.ReplaceString(DITEM.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREPLAYLAN.Text) + "', "


        'QCNO,ACCDEP1,ACCDEP2,COST,ACCEMPNAME,JUDGE1,REMARK1,CPNO,HAPPEN1,CONFIRM1,MANUFLOW,REMARK2,YKKAGE,HAPPEN2
        sql = sql + " '', "
        sql = sql + " N'" + YKK.ReplaceString(DQCNO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCDEP1.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCDEP12.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCDEP13.SelectedValue) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DTYPE.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DPL.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCEMPNAME.Text) + "', "


        sql = sql + " N'" + YKK.ReplaceString(DACCDEP2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DACCDEP2NO.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCUSTOMERTYPE.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DJUDGE.SelectedValue) + "', "

        sql = sql + " '', "
        sql = sql + " N'" + YKK.ReplaceString(DMANUREASON.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DMANUREASON1.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DRESPONS.SelectedValue) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DHOUR1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DHOUR2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DHOUR3.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DHOUR4.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DHOUR5.Text) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DSAMPLEDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DANSWERDAYS.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK.SelectedValue) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DACOST.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DBCOST.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCCOST.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DDCOST.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DECOST.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DALLCOST.Text) + "', "
        'REMARK3,POINT1,POINT2,POINT3
        sql = sql + " N'" + YKK.ReplaceString(DPOINT1.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DPOINT2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DPOINT3.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCOST.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCUSTDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DYFGCDATE.Text) + "', "
        sql = sql + " N'" + WORKPAY + "', "
        sql = sql + " '0', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREMARK3.Text) + "', "

        sql = sql + " N'" + YKK.ReplaceString(DSCORE1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSCORE2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSCORE3.Text) + "', "

        If DRELY.Checked = True Then
            sql = sql + " 1, "
        Else
            sql = sql + " 0, "
        End If


        If DMAT.Checked = True Then
            sql = sql + " 1, "
        Else
            sql = sql + " 0, "
        End If


        sql = sql + " N'" + YKK.ReplaceString(DQCCDESC.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCEDESC.Text) + "', "

        'sql = sql + " DELIVERYDATE,SAMPLE,TYPE1,TYPE2,ReportDate,QCDESCTYPE,JUDGE1,JUDGE2,CARDATE,FIRSTDATE,MIDDATE,"

        sql = sql + " N'" + YKK.ReplaceString(DDELIVERYDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DSAMPLE.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DTYPE1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DTYPE2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DREPORTDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCDESCTYPE.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DJUDGE1.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DJUDGE2.SelectedValue) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DCARDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DFIRSTDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DFINALDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DMIDDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DQCACCDATE.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DMFGDATE.Text) + "', "


        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)


        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '3109'"
        sql = sql + " and dkey ='AttachfilePath' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter1.Rows.Count > 0 Then
            sourceDir = "\\" + DBAdapter1.Rows(0).Item("Data1") + D3.Text   '來源
        End If

        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '3109'"
        sql = sql + " and dkey ='AttachfilePath1' "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter2.Rows.Count > 0 Then
            backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNO.Text      '目的  
        End If


        'sourceDir = "\\10.245.1.61\wfs$\N2W\003109Temp\" + D3.Text   '來源
        'backupDir = "\\10.245.1.61\wfs$\N2W\003109\" + DNO.Text      '目的     

        '複製前先確認是否有原本的資料夾再刪除
        If Directory.Exists(backupDir) Then
            Directory.Delete(backupDir, True)
        End If
        CopyDir(sourceDir, backupDir)
        NewAttachFilePath()



    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                                         CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                  System.Configuration.ConfigurationManager.AppSettings("ComplaintOutPath")
        Dim FileName As String = ""
        Dim WORKPAY As String = ""
        Dim sql1 As String
        Dim i As Integer

        If wStep = 10 Then  '勞務費以品管受理日為基準
            SQL1 = " SELECT * FROM    m_referp"
            SQL1 = SQL1 + "  where cat = '3109'"
            SQL1 = SQL1 + " and dkey = 'WORKPAY'"
            Dim dtReferp As DataTable = uDataBase.GetDataTable(sql1)
            For i = 0 To dtReferp.Rows.Count - 1
                WORKPAY = dtReferp.Rows(i).Item("Data")
                DWORKPAY.Text = WORKPAY
            Next
            dtReferp.Clear()
        End If

        Dim sql As String
        sql = " Update F_ComplaintOutSheet"
        sql = sql + " Set "
        If pFun <> "SAVE" Then
            sql = sql + " Sts = '" & pSts & "',"
            sql = sql + " CompletedTime = '" & NowDateTime & "',"
        End If
        sql = sql + " No = N'" & YKK.ReplaceString(DNO.Text) & "',"
        sql = sql + " DATE = N'" + DDATE.Text + "', "   'NO  1


        sql = sql + " GLOBAL = N'" + YKK.ReplaceString(DGLOBAL.SelectedValue) + "', "
        sql = sql + " CUSTOMERCODE = N'" + YKK.ReplaceString(DCUSTOMERCODE.Text) + "', "
        sql = sql + " CUSTOMER = '" + YKK.ReplaceString(DCUSTOMER.Text) + "', "
        sql = sql + " APPNAME = N'" + YKK.ReplaceString(DAPPNAME.SelectedValue) + "', "

        sql = sql + " OORNO = '" + YKK.ReplaceString(DOORNO.Text) + "', "
        sql = sql + " NORNO = '" + YKK.ReplaceString(DNORNO.Text) + "', "
        sql = sql + " SPEC = '" + YKK.ReplaceString(DSPEC.Text) + "', "
        sql = sql + " SPECNAME = '" + YKK.ReplaceString(DSPECNAME.Text) + "', "
        sql = sql + " ORQTY = '" + YKK.ReplaceString(DORQTY.Text) + "', "

        sql = sql + " ORDERDATE = '" + YKK.ReplaceString(DORDERDATE.Text) + "', "
        sql = sql + " SHIPDATE= '" + YKK.ReplaceString(DSHIPDATE.Text) + "', "


        sql = sql + " BIGGOODS = N'" + YKK.ReplaceString(DBIGGOODS.SelectedValue) + "', "
        sql = sql + " PLACE = N'" + YKK.ReplaceString(DPLACE.SelectedValue) + "', "
        sql = sql + " CPCONTENT= N'" + YKK.ReplaceString(DCPCONTENT.Text) + "', "
        sql = sql + " EPCONTENT= N'" + YKK.ReplaceString(DEPCONTENT.Text) + "', "
        sql = sql + " CPQTY= N'" + YKK.ReplaceString(DCPQTY.Text) + "', "
        sql = sql + " WORKPAY='" + DWORKPAY.Text + "', " '勞務費

        If DCPQTYCHK.Checked Then
            sql = sql + " CPQTYCHK= 1, "
        Else
            sql = sql + " CPQTYCHK= 0, "


        End If

        sql = sql + " BuyerCode= N'" + YKK.ReplaceString(DBUYERCODE.Text) + "', "
        sql = sql + " Buyer= N'" + YKK.ReplaceString(DBUYER.Text) + "', "


        FileName = ""


        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                'FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & System.IO.Path.GetExtension(DMapFile.PostedFile.FileName)
                DMapFile.PostedFile.SaveAs(Path + FileName)
                sql = sql + " MapFile= N'" + YKK.ReplaceString(FileName) + "',"           '不良樣品圖
            Else
                FileName = ""
            End If
        End If







        sql = sql + " ITEM = N'" + YKK.ReplaceString(DITEM.SelectedValue) + "', "
        sql = sql + " REPLAYLAN = N'" + YKK.ReplaceString(DREPLAYLAN.SelectedValue) + "', "

        If wStep = 10 Then
            sql = sql + " QCRDATE= N'" + Now.ToString("yyyy/MM/dd") + "', "
        End If
        sql = sql + " QCNO= N'" + YKK.ReplaceString(DQCNO.Text) + "', "
        sql = sql + " ACCDEP1= N'" + YKK.ReplaceString(DACCDEP1.SelectedValue) + "', "
        sql = sql + " ACCDEP12= N'" + YKK.ReplaceString(DACCDEP12.SelectedValue) + "', "
        sql = sql + " ACCDEP13= N'" + YKK.ReplaceString(DACCDEP13.SelectedValue) + "', "

        sql = sql + " TYPE= N'" + YKK.ReplaceString(DTYPE.SelectedValue) + "', "
        sql = sql + " PL= N'" + YKK.ReplaceString(DPL.SelectedValue) + "', "
        sql = sql + " ACCEMPNAME= N'" + YKK.ReplaceString(DACCEMPNAME.Text) + "', "


        sql = sql + " ACCDEP2= N'" + YKK.ReplaceString(DACCDEP2.SelectedValue) + "', "
        sql = sql + " ACCDEP2NO= N'" + YKK.ReplaceString(DACCDEP2NO.Text) + "', "
        sql = sql + " CUSTOMERTYPE= N'" + YKK.ReplaceString(DCUSTOMERTYPE.SelectedValue) + "', "
        sql = sql + " JUDGE= N'" + YKK.ReplaceString(DJUDGE.Text) + "', "

        If wStep = 20 Then
            sql = sql + " MANURDATE= N'" + Now.ToString("yyyy/MM/dd") + "', "
        End If
        sql = sql + " MANUREASON = N'" + YKK.ReplaceString(DMANUREASON.SelectedValue) + "', "
        sql = sql + " MANUREASON1 = N'" + YKK.ReplaceString(DMANUREASON1.SelectedValue) + "', "
        sql = sql + " RESPONS = N'" + YKK.ReplaceString(DRESPONS.SelectedValue) + "', "

        sql = sql + " HOUR1= N'" + YKK.ReplaceString(DHOUR1.Text) + "', "
        sql = sql + " HOUR2= N'" + YKK.ReplaceString(DHOUR2.Text) + "', "
        sql = sql + " HOUR3= N'" + YKK.ReplaceString(DHOUR3.Text) + "', "
        sql = sql + " HOUR4= N'" + YKK.ReplaceString(DHOUR3.Text) + "', "
        sql = sql + " HOUR5= N'" + YKK.ReplaceString(DHOUR3.Text) + "', "

        sql = sql + " SAMPLEDATE = N'" + YKK.ReplaceString(DSAMPLEDATE.Text) + "', "
        sql = sql + " QCDATE = N'" + YKK.ReplaceString(DQCDATE.Text) + "', "
        sql = sql + " ANSWERDAYS = N'" + YKK.ReplaceString(DANSWERDAYS.Text) + "', "

        sql = sql + " ACOST = N'" + YKK.ReplaceString(DACOST.Text) + "', "
        sql = sql + " BCOST = N'" + YKK.ReplaceString(DBCOST.Text) + "', "
        sql = sql + " CCOST = N'" + YKK.ReplaceString(DCCOST.Text) + "', "

        If DHOUR1.Text <> "" And DHOUR2.Text <> "" And DHOUR5.Text <> "" Then
            sql = sql + " DCOST =  convert(float,workpay)/8*" + CStr(CInt(DHOUR1.Text) + CInt(DHOUR2.Text) + CInt(DHOUR3.Text) + CInt(DHOUR4.Text)) + "+" + YKK.ReplaceString(DHOUR5.Text) + ", "
        End If

        sql = sql + " ECOST = N'" + YKK.ReplaceString(DECOST.Text) + "', "
        sql = sql + " FCOST = N'" + YKK.ReplaceString(DFCOST.Text) + "', "
        sql = sql + " GCOST = N'" + YKK.ReplaceString(DGCOST.Text) + "', "
        sql = sql + " HCOST = N'" + YKK.ReplaceString(DHCOST.Text) + "', "

        sql = sql + " ALLCOST = N'" + YKK.ReplaceString(DALLCOST.Text) + "', "

        sql = sql + " POINT1 = N'" + Trim(YKK.ReplaceString(DPOINT1.SelectedValue)) + "', "
        sql = sql + " POINT2 = N'" + YKK.ReplaceString(DPOINT2.Text) + "', "
        sql = sql + " POINT3 = N'" + YKK.ReplaceString(DPOINT3.Text) + "', "


        sql = sql + " SCORE1 = N'" + YKK.ReplaceString(DSCORE1.SelectedValue) + "', "
        sql = sql + " SCORE2 = N'" + YKK.ReplaceString(DSCORE2.SelectedValue) + "', "
        sql = sql + " SCORE3 = N'" + YKK.ReplaceString(DSCORE3.SelectedValue) + "', "

        sql = sql + " COST= N'" + YKK.ReplaceString(DCOST.Text) + "', "

        sql = sql + " CUSTDATE = N'" + YKK.ReplaceString(DCUSTDATE.Text) + "', "
        sql = sql + " YFGCDATE = N'" + YKK.ReplaceString(DYFGCDATE.Text) + "', "
        If wStep = 30 Then 'QC最終確認
            sql = sql + " QCSts =1 , "
        Else
            sql = sql + " QCSts =0 , "
        End If
        sql = sql + " Remark1 = N'" + YKK.ReplaceString(DREMARK1.Text) + "', "
        sql = sql + " Remark2 = N'" + YKK.ReplaceString(DREMARK2.Text) + "', "
        sql = sql + " Remark3 = N'" + YKK.ReplaceString(DREMARK3.Text) + "', "
        If DRELY.Checked = True Then
            sql = sql + " RELY = 1, "

        Else
            sql = sql + " RELY = 0, "
        End If

        If DMAT.Checked = True Then
            sql = sql + " MAT = 1, "

        Else
            sql = sql + " MAT = 0, "

        End If

        If DRELY.Checked = True Or DMAT.Checked = True Then
            If DRELY.Checked Then
                sql = sql + " REMARK='*分析依賴',"
            Else
                sql = sql + " REMARK='*特料抱怨單',"
            End If
        Else
            sql = sql + " REMARK = N'" + YKK.ReplaceString(DREMARK.Text) + "', "

        End If

        sql = sql + " QCCDESC = N'" + YKK.ReplaceString(DQCCDESC.Text) + "', "
        sql = sql + " QCEDESC = N'" + YKK.ReplaceString(DQCEDESC.Text) + "', "


        If DMach1.Checked = True Then
            sql = sql + " MACH='0',"
            sql = sql + " MACHNO= N'" + YKK.ReplaceString(DMACHNO.Text) + "', "
        ElseIf DMach2.Checked = True Then
            sql = sql + " MACH='1',"
            sql = sql + " MACHNO= '', "
        Else
            sql = sql + " MACH='',"
            sql = sql + " MACHNO= '', "
        End If
        sql = sql + " HAPPEN='" + DHappen.SelectedValue + "',"


        'sql = sql + " DELIVERYDATE,SAMPLE,TYPE1,TYPE2,ReportDate,QCDESCTYPE,JUDGE1,JUDGE2,CARDATE,FIRSTDATE,MIDDATE,"

        sql = sql + " DELIVERYDATE='" + DDELIVERYDATE.Text + "',"
        sql = sql + " SAMPLE='" + DSAMPLE.SelectedValue + "',"
        sql = sql + " TYPE1='" + DTYPE1.SelectedValue + "',"
        sql = sql + " TYPE2='" + DTYPE2.SelectedValue + "',"
        sql = sql + " ReportDate='" + DREPORTDATE.Text + "',"
        sql = sql + " QCDESCTYPE='" + DQCDESCTYPE.SelectedValue + "',"
        sql = sql + " JUDGE1='" + DJUDGE1.SelectedValue + "',"
        sql = sql + " JUDGE2='" + DJUDGE2.SelectedValue + "',"
        sql = sql + " CARDATE='" + DCARDATE.Text + "',"
        sql = sql + " FIRSTDATE='" + DFIRSTDATE.Text + "',"
        sql = sql + " FINALDATE='" + DFINALDATE.Text + "',"
        sql = sql + " MIDDATE='" + DMIDDATE.Text + "',"
        sql = sql + " QCACCDATE='" + DQCACCDATE.Text + "',"
        sql = sql + " MFGDATE='" + DMFGDATE.Text + "',"

        sql = sql + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql = sql + " ModifyTime = '" & NowDateTime & "' "
        sql = sql + " Where FormNo  =  '" & wFormNo & "'"
        sql = sql + "   And FormSno =  '" & CStr(wFormSno) & "'"
        '
        uDataBase.ExecuteNonQuery(sql)
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
        Dim dtCommissionNo As DataTable = uDataBase.GetDataTable(SQl)

        If dtCommissionNo.Rows.Count <= 0 Then
            If DNO.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNO.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNO.Text <> "" Then
                If DNO.Text <> dtCommissionNo.Rows(0)("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNO.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
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
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim SQl As String
        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Text & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Text & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Text & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            'jessica 2021/1/4 退回關卡同時有一人以上
            ' SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
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
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BSAVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.Click
        ' If Request.Cookies("RunBSAVE").Value = True Then

        Try
            ModifyData("SAVE", "0")
            ModifyTranData("SAVE", "0")
            uJavaScript.PopMsg(Me, "存檔成功")
        Catch ex As Exception
            uJavaScript.PopMsg(Me, "存檔失敗請洽電腦室")
        End Try

        'End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG2.Click
        DisabledButton()   '停止Button運作
        FlowControl("NG2", 2, "3")
        'If InputDataOK(2) Then
        '    DisabledButton()   '停止Button運作
        '    FlowControl("NG2", 2, "3")
        'Else
        '    EnabledButton()   '起動Button運作
        'End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG1.Click
        DisabledButton()   '停止Button運作
        FlowControl("NG1", 1, "2")

        'If InputDataOK(1) Then
        '    DisabledButton()   '停止Button運作
        '    FlowControl("NG1", 1, "2")
        'Else
        '    EnabledButton()   '起動Button運作
        'End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Enabled = False
        BNG1.Enabled = False
        BNG2.Enabled = False
        BSAVE.Enabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Enabled = True
        BNG1.Enabled = True
        BNG2.Enabled = True
        BSAVE.Enabled = True
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
        Dim Str3 As String = ""
        Dim i As Integer



        'Set當日日期
        Str3 = Mid(CStr(DateTime.Now.Year), 3, 2)  '年

        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str3 + Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 4 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function
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
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9020
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9030
            End If
        End If
    End Function

    '*****************************************************************
    '**
    '**     判斷是否可繼續執行(驗證資料)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim acost, bcost, ccost, dcost, ecost, allcost As Integer
        Dim sql As String

        If ErrCode = 0 Then
            If Len(DCPCONTENT.Text) > 250 Then
                ErrCode = 9089
            End If
        End If

        If ErrCode = 0 Then
            If Len(DEPCONTENT.Text) > 250 Then
                ErrCode = 9089
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

        'Check品管必輸入 
        If ErrCode = 0 Then
            If DQCNO.Visible = True Then
                If DQCNO.Text = "" Then
                    ErrCode = 9061
                End If
            End If
        End If
        If ErrCode = 0 Then
            If DTYPE.Visible = True Then
                If DTYPE.SelectedValue = "" Then
                    ErrCode = 9062
                End If
            End If
        End If

        If wStep = 10 Then
            If ErrCode = 0 Then
                If DACCDEP1.Visible = True Then
                    If DACCDEP1.SelectedValue = "" Then
                        ErrCode = 9063
                    End If
                    If DTYPE.SelectedValue <> "依賴" Then
                        If DACCDEP12.SelectedValue = "" Then
                            ErrCode = 9086
                        End If
                        If DACCDEP13.SelectedValue = "" Then
                            ErrCode = 9087
                        End If
                    End If

                End If
            End If
            If ErrCode = 0 Then
                If DCUSTOMERTYPE.Visible = True Then
                    If DCUSTOMERTYPE.SelectedValue = "" Then
                        ErrCode = 9064
                    End If
                End If
            End If

            If DTYPE.SelectedValue <> "依賴" Then
                If ErrCode = 0 Then
                    If DQCDESCTYPE.SelectedValue = "" Then
                        ErrCode = 9088
                    End If
                End If


            End If

        End If
       



        '檢查資料夾是否有檔案

        If wStep = 1 Or wStep = 500 Then


            If ErrCode = 0 Then
                Dim dirInfo As New System.IO.DirectoryInfo(chktemp.Text)
                Dim FileDir As Integer  '資料夾
                FileDir = dirInfo.GetDirectories("*").Length
                Dim FileCount As Integer '檔案
                FileCount = dirInfo.GetFiles("*.*").Length
                If FileCount > 0 Or FileDir > 0 Then
                    DChkData1.Checked = True

                Else
                    DChkData1.Checked = False
                    ErrCode = 9065
                End If
            End If

        End If

        If wStep = 20 Then
            If ErrCode = 0 Then
                If DACCDEP2.Visible = True Then
                    If DACCDEP2.SelectedValue = "" Then
                        ErrCode = 9066
                    End If
                End If
            End If


            If ErrCode = 0 Then
                If DJUDGE.Visible = True Then
                    If DJUDGE.SelectedValue = "" Then
                        ErrCode = 9067
                    End If
                End If
            End If


            If ErrCode = 0 Then
                If DMANUREASON.Visible = True Then
                    If DMANUREASON.SelectedValue = "" Then
                        ErrCode = 9068
                    End If
                End If
            End If



            If ErrCode = 0 Then
                If DMANUREASON1.Visible = True Then
                    If DMANUREASON1.SelectedValue = "" Then
                        ErrCode = 9069
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DRESPONS.Visible = True Then
                    If DRESPONS.SelectedValue = "" Then
                        ErrCode = 9070
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DHOUR1.Visible = True Then
                    If DHOUR1.Text = "" Then
                        ErrCode = 9071
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DHOUR2.Visible = True Then
                    If DHOUR2.Text = "" Then
                        ErrCode = 9072
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DHOUR3.Visible = True Then
                    If DHOUR3.Text = "" Then
                        ErrCode = 9073
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DHOUR4.Visible = True Then
                    If DHOUR4.Text = "" Then
                        ErrCode = 9074
                    End If
                End If
            End If


        End If


        If wStep = 35 Then
            If ErrCode = 0 Then
                If DACOST.Visible = True Then
                    If DACOST.Text = "" Then
                        ErrCode = 9075
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DBCOST.Visible = True Then
                    If DBCOST.Text = "" Then
                        ErrCode = 9076
                    End If
                End If
            End If


            If ErrCode = 0 Then
                If DCCOST.Visible = True Then
                    If DCCOST.Text = "" Then
                        ErrCode = 9077
                    End If
                End If
            End If


            If ErrCode = 0 Then
                If DECOST.Visible = True Then
                    If DECOST.Text = "" Then
                        ErrCode = 9078
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DCUSTDATE.Visible = True Then
                    If DCUSTDATE.Text = "" Then
                        ErrCode = 9081
                    End If
                End If
            End If

            If ErrCode = 0 Then
                If DYFGCDATE.Visible = True Then
                    If DYFGCDATE.Text = "" Then
                        ErrCode = 9080
                    End If
                End If
            End If
        End If

        '機台號碼
        If ErrCode = 0 Then
            If DMach1.Checked = True And DMach2.Checked Then
                ErrCode = 9084
            ElseIf DMach1.Checked = True Then
                If DMACHNO.Text = "" Then
                    ErrCode = 9083
                End If
            End If

        End If


        If wStep = 20 Then
            If ErrCode = 0 Then
                If DMach1.Checked = False And DMach2.Checked = False Then
                    ErrCode = 9085
                End If
            End If
        End If




        'Check上傳附件Size及格式
        If ErrCode = 0 Then
            If DMapFile.Visible Then
                If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DMapFile)
                End If
            End If
        End If




        '檢查委託書No
        If ErrCode = 0 Then
            If DNO.Text <> "" Then
                ErrCode = oCommon.CommissionNo("003109", wFormSno, wStep, DNO.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If

        '檢查QCNO 是否重覆
        If wStep = 10 Then
            If ErrCode = 0 Then

                sql = " select  * from F_ComplaintOutSheet   "
                sql = sql + "where qcno ='" + DQCNO.Text + "'"
                sql = sql + " and no ='" + DNO.Text + "'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)
                If DBAdapter1.Rows.Count > 1 Then
                    ErrCode = 9074
                End If


            End If

        End If

    

        Try

            If ErrCode = 0 Then

                If DHOUR1.Text <> "" Then
                    DHOUR1.Text = CInt(DHOUR1.Text)
                End If
                If DHOUR2.Text <> "" Then
                    DHOUR2.Text = CInt(DHOUR2.Text)
                End If
                If DHOUR3.Text <> "" Then
                    DHOUR3.Text = CInt(DHOUR3.Text)
                End If

                If DHOUR3.Text <> "" Then
                    DHOUR3.Text = CInt(DHOUR3.Text)
                End If

                If DHOUR4.Text <> "" Then
                    DHOUR4.Text = CInt(DHOUR4.Text)
                End If

                If DHOUR5.Text <> "" Then
                    DHOUR5.Text = CInt(DHOUR5.Text)
                End If

                If DPOINT2.Text <> "" Then
                    DPOINT3.Text = CInt(DPOINT2.Text)
                End If


                If DPOINT1.SelectedValue <> "" Then
                    DPOINT3.Text = CInt(DPOINT1.SelectedValue) + CInt(DPOINT2.Text)
                End If


                If DACOST.Text <> "" Then
                    acost = CInt(DACOST.Text)
                End If
                If DBCOST.Text <> "" Then
                    bcost = CInt(DBCOST.Text)
                End If
                If DCCOST.Text <> "" Then
                    ccost = CInt(DCCOST.Text)
                End If
                If DDCOST.Text <> "" Then
                    dcost = CInt(DDCOST.Text)
                End If
                If DECOST.Text <> "" Then
                    ecost = CInt(DECOST.Text)
                End If

                allcost = acost + bcost + ccost + dcost + ecost
                DALLCOST.Text = Str(allcost)


            End If

        Catch ex As Exception
            ErrCode = 9087
        End Try

        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "Item資料有誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認是否為JPG圖檔!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"

            If ErrCode = 9061 Then Message = "請輸入品管受理編號!"
            If ErrCode = 9062 Then Message = "請輸入客訴判定!"
            If ErrCode = 9063 Then Message = "請輸入責任部門別!"
            If ErrCode = 9064 Then Message = "請輸入顧客!"
            If ErrCode = 9065 Then Message = "請開啟資料夾放入營業登錄附件!"
            If ErrCode = 9066 Then Message = "請輸入責任工程別名稱!"
            If ErrCode = 9067 Then Message = "請輸入判定!"
            If ErrCode = 9068 Then Message = "請輸入原因!"
            If ErrCode = 9069 Then Message = "請輸入人為錯誤分類!"
            If ErrCode = 9070 Then Message = "請輸入責任區分/起因!"
            If ErrCode = 9071 Then Message = "請輸入調查處理工時!"
            If ErrCode = 9072 Then Message = "請輸入檢查工時!"
            If ErrCode = 9073 Then Message = "請輸入國內出差工時!"
            If ErrCode = 9074 Then Message = "請輸入客訴品修理工時!"
            If ErrCode = 9075 Then Message = "請輸入賠償費用（a)!"
            If ErrCode = 9076 Then Message = "請輸入運費（b)!"
            If ErrCode = 9077 Then Message = "請輸入出差費用（c)!"
            If ErrCode = 9078 Then Message = "請輸入委外測費用（e)!"
            If ErrCode = 9079 Then Message = "請輸入重要度!"
            If ErrCode = 9080 Then Message = "請輸入YFGC登錄日!"
            If ErrCode = 9081 Then Message = "請輸入顧客說明日!"
            If ErrCode = 9082 Then Message = "非數字格式,請確認!"
            If ErrCode = 9083 Then Message = "請輸入該工程POP機台號!"
            If ErrCode = 9084 Then Message = "只能勾選機台或非機台其中一項!"
            If ErrCode = 9085 Then Message = "請勾選機台或非機台其中一項!"
            If ErrCode = 9086 Then Message = "請輸入Where-不良位置!"
            If ErrCode = 9087 Then Message = "請輸入What-什麼問題!"
            If ErrCode = 9088 Then Message = "請輸入客訴內容分類!"
            If ErrCode = 9089 Then Message = "中文客訴內容字數不可超過250字元，請再修正!"
            If ErrCode = 9090 Then Message = "英文客訴內容字數不可超過250字元，請再修正!"

            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        Return isOK
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑 1
    '**
    '*****************************************************************
    Sub NewAttachFilePath()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/營業"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\營業"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        chktemp.Text = tempFolderPath

        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑 2
    '**
    '*****************************************************************
    Sub NewAttachFilePath2()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNO.Text + "/品管確認"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\品管確認"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length

        If FileCount > 0 Or FileDir > 0 Then
            DChkData2.Checked = True
            '   DChkData2.Text = Str(FileCount) + "件"
        Else
            DChkData2.Checked = False
            ' DChkData2.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath3()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNO.Text + "/製造"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\製造"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length

        If FileCount > 0 Or FileDir > 0 Then
            DChkData3.Checked = True
            ' DChkData3.Text = Str(FileCount) + "件"
        Else
            DChkData3.Checked = False
            DChkData3.Text = ""
        End If

        '開啟附檔資料夾路徑
        DAttachfile3.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath4()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + DNO.Text + "/品管最終確認"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\品管最終確認"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length


        If FileCount > 0 Or FileDir > 0 Then
            DChkData4.Checked = True
            ' DChkData4.Text = Str(FileCount) + "件"
        Else
            DChkData4.Checked = False
            ' DChkData4.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile4.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath5()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + DNO.Text + "/未然防止T"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\未然防止T"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length

        If FileCount > 0 Or FileDir > 0 Then
            DChkData5.Checked = True
            ' DChkData5.Text = Str(FileCount) + "件"
        Else
            DChkData5.Checked = False
            ' DChkData5.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile5.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath6()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + DNO.Text + "/CAR簽核版"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\CAR簽核版"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)
        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length

        If FileCount > 0 Or FileDir > 0 Then
            DChkData6.Checked = True
            ' DChkData5.Text = Str(FileCount) + "件"
        Else
            DChkData6.Checked = False
            ' DChkData5.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile6.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub



    '複製資料夾
    Public Shared Sub CopyDir(ByVal srcPath As String, ByVal aimPath As String)
        Try
            ' 檢查目標目錄是否以目錄分割字元結束如果不是則新增之
            If aimPath(aimPath.Length - 1) <> Path.DirectorySeparatorChar Then
                aimPath += Path.DirectorySeparatorChar
            End If
            ' 判斷目標目錄是否存在如果不存在則新建之
            If Not Directory.Exists(aimPath) Then
                Directory.CreateDirectory(aimPath)
            End If
            ' 得到源目錄的檔案列表，該裡面是包含檔案以及目錄路徑的一個數組
            ' 如果你指向copy目標檔案下面的檔案而不包含目錄請使用下面的方法
            ' string[] fileList = Directory.GetFiles(srcPath);
            Dim fileList() As String = Directory.GetFileSystemEntries(srcPath)
            ' 遍歷所有的檔案和目錄
            Dim file As String
            For Each file In fileList
                ' 先當作目錄處理如果存在這個目錄就遞迴Copy該目錄下面的檔案
                If Directory.Exists(file) Then
                    CopyDir(file, aimPath + Path.GetFileName(file))
                Else
                    System.IO.File.Copy(file, aimPath + Path.GetFileName(file), True)
                End If

                ' 否則直接Copy檔案

            Next
        Catch
            Console.WriteLine("無法複製!")
        End Try
    End Sub

    '取得關係人
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID,RRUserID,USERID  from M_Related where userid='" & userId & "' and RelatedID='A'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            If wStep = 30 Then
                If dt.Rows(0)("UserID") = dt.Rows(0)("RUserID") Then '如果被訴擔當跟被訴擔當台籍主管一樣就直接跳日籍主管
                    NextGate = dt.Rows(0)("RRUserID")
                Else
                    NextGate = dt.Rows(0)("RUserID")
                End If

            ElseIf wStep = 40 Then
                NextGate = dt.Rows(0)("RRUserID")
            End If

        End If
        Return NextGate
    End Function

    '取得關係人
    Function GetUerID(ByVal Username As String) As String

        Dim sql As String = "select UserID from M_users where "
        sql = sql + "  username=N'" + Trim(Username) + "'"
        sql = sql + " and pluralism =0"

        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserID")
        End If
        Return NextGate
    End Function

    '取得關係人
    Function GetUerNname(ByVal UserID As String) As String

        Dim sql As String = "select UserName from M_users where userid='" + Trim(UserID) + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserName")
        End If
        Return NextGate
    End Function


    Protected Sub DACCDEP1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACCDEP1.SelectedIndexChanged
        Dim sql As String
        Dim username As String = ""
        sql = " Select SUBSTRING(DATA,CHARINDEX('-',DATA)+1,LEN(DATA)-1)DATA  from M_referp"
        sql = sql & " where  cat = '3109'"
        sql = sql & " and dkey = 'RDIVISION'"
        sql = sql & " and substring(data,1,CHARINDEX('-', data)-1)='" + Trim(DACCDEP1.SelectedValue) + "'"
        sql = sql & " order by unique_id"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            username = dt.Rows(0)("data")
            DACCEMPNAME.Text = dt.Rows(0)("data")
        Else
            DACCEMPNAME.Text = ""
        End If

        LMapFile.ImageUrl = DMappath.Text
        LMapFile1.NavigateUrl = DMappath.Text
        LMapFile.Visible = True
        LMapFile1.Visible = True



        Dim i As Integer
     
        sql = "  select data,substring(dkey,8,len(dkey)-1) from  M_referp"
        sql = sql & " where  left(Dkey,6)='WhereP'"
        sql = sql & " and substring(dkey,8,len(dkey)-1) in("
        sql = sql & " select  substring(dkey,10,len(dkey)-1)Data from  M_referp where cat ='3109' and left(Dkey,8) ='WhereDep'"
        sql = sql & " and data like '%" + Trim(DACCDEP1.SelectedValue) + "%') order by Unique_id"

        Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
        DACCDEP12.Items.Clear()
        DACCDEP12.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")

            DACCDEP12.Items.Add(ListItem1)
        Next




    End Sub

    '取得關係人
    Function GetRelated1(ByVal userId As String) As String

        Dim sql As String = "select RUserID,RRUserID,USERID  from M_Related where userid='" & userId & "' and RelatedID='A'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            If dt.Rows(0)("UserID") = dt.Rows(0)("RUserID") Then '如果被訴擔當跟被訴擔當台籍主管一樣就直接跳日籍主管
                NextGate = dt.Rows(0)("RRUserID")
            Else
                NextGate = dt.Rows(0)("RUserID")
            End If

        End If
        Return NextGate
    End Function

    Sub Total()
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim acost, bcost, ccost, dcost, ecost, fcost, gcost, hcost, allcost, HOUR, d1, d2, d3, d4, d5 As Integer
        Dim AWORKPAY As Decimal

        Try

            AWORKPAY = CInt(DWORKPAY.Text) / 8
            '處理費用 =(D1+D2+D3+D4)*人工薪+D5
            'If DHOUR1.Text <> "" And DHOUR2.Text <> "" And DHOUR3.Text <> "" And DHOUR4.Text <> "" And DHOUR5.Text <> "" Then
            If DHOUR1.Text = "" Then
                d1 = 0
            Else
                d1 = DHOUR1.Text
            End If
            If DHOUR2.Text = "" Then
                d2 = 0
            Else
                d2 = DHOUR2.Text
            End If
            If DHOUR3.Text = "" Then
                d3 = 0
            Else
                d3 = DHOUR3.Text
            End If
            If DHOUR4.Text = "" Then
                d4 = 0
            Else
                d4 = DHOUR4.Text
            End If
            If DHOUR5.Text = "" Then
                d5 = 0
            Else
                d5 = DHOUR5.Text
            End If

            HOUR = (CInt(d1) + CInt(d2) + CInt(d3) + CInt(d4)) * AWORKPAY + CInt(d5)
            DDCOST.Text = HOUR

            If DPOINT2.Text <> "" Then
                DPOINT3.Text = CInt(DPOINT2.Text)
            End If

            If DPOINT1.SelectedValue <> "" Then
                DPOINT3.Text = CInt(DPOINT1.SelectedValue) + CInt(DPOINT2.Text)
            End If

            'End If




            If DACOST.Text <> "" Then
                acost = CInt(DACOST.Text)
            End If
            If DBCOST.Text <> "" Then
                bcost = CInt(DBCOST.Text)
            End If
            If DCCOST.Text <> "" Then
                ccost = CInt(DCCOST.Text)
            End If
            If DDCOST.Text <> "" Then
                dcost = CInt(DDCOST.Text)
            End If
            If DECOST.Text <> "" Then
                ecost = CInt(DECOST.Text)
            End If
            If DFCOST.Text <> "" Then
                fcost = CInt(DFCOST.Text)
            End If

            If DGCOST.Text <> "" Then
                gcost = CInt(DGCOST.Text)
            End If

            If DHCOST.Text <> "" Then
                hcost = CInt(DHCOST.Text)
            End If



            allcost = acost + bcost + ccost + dcost + ecost + fcost + gcost + hcost
            DALLCOST.Text = Str(allcost)

            If CInt(DALLCOST.Text) <= 20500 Then
                DPOINT2.Text = "0"
            ElseIf CInt(DALLCOST.Text) > 20501 And CInt(DALLCOST.Text) <= 205000 Then
                DPOINT2.Text = "1"
            ElseIf CInt(DALLCOST.Text) > 205001 And CInt(DALLCOST.Text) <= 1720500 Then
                DPOINT2.Text = "3"
            ElseIf CInt(DALLCOST.Text) > 1720501 And CInt(DALLCOST.Text) <= 2050000 Then
                DPOINT2.Text = "5"
            ElseIf CInt(DALLCOST.Text) > 2050001 Then
                DPOINT2.Text = "10"
            End If


            LMapFile.ImageUrl = DMappath.Text
            LMapFile1.NavigateUrl = DMappath.Text
            LMapFile.Visible = True
            LMapFile1.Visible = True

        Catch ex As Exception
            uJavaScript.PopMsg(Me, "請輸入數字")
        End Try

    End Sub

    Protected Sub DACOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DACOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DBCOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBCOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DBCOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DECOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DECOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DECOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DPOINT3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Total()

    End Sub

    Protected Sub DPOINT1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DPOINT1.SelectedIndexChanged
        Total()
    End Sub

    Protected Sub DHOUR1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DHOUR1.TextChanged
    Total()

    End Sub

    Protected Sub DHOUR2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DHOUR2.TextChanged
       Total()
    End Sub

    Protected Sub DHOUR3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DHOUR3.TextChanged
    Total()
    End Sub

    Protected Sub DHOUR4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DHOUR4.TextChanged
       Total()
    End Sub

    Protected Sub DHOUR5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DHOUR5.TextChanged
       Total()
    End Sub
    
    Protected Sub DMANUREASON_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMANUREASON.SelectedIndexChanged
        If DMANUREASON.SelectedValue = "1.人為" Then
            Dim sql As String
            Dim i As Integer
            Sql = "  Select DATA from M_referp"
            Sql = Sql & " where  cat = '3109'"
            Sql = Sql & " and dkey = 'MANUREASON1'"
            Sql = Sql & " order by unique_id"

            Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
            DMANUREASON1.Items.Clear()
            DMANUREASON1.Items.Add("")
            For i = 0 To dtReferp.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = dtReferp.Rows(i).Item("Data")
                ListItem1.Value = dtReferp.Rows(i).Item("Data")
                DMANUREASON1.Items.Add(ListItem1)
            Next
            dtReferp.Clear()
        Else
            DMANUREASON1.Items.Clear()
            DMANUREASON1.Items.Add("無")
        End If

        LMapFile.ImageUrl = DMappath.Text
        LMapFile1.NavigateUrl = DMappath.Text
        LMapFile.Visible = True
        LMapFile1.Visible = True
        
    End Sub

    Protected Sub DACCDEP2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACCDEP2.SelectedIndexChanged
        Dim sql As String
        Sql = "  Select substring(data,9,len(data)-1) data1,substring(data,1,7) data2  from M_referp"
        Sql = Sql & " where  cat = '3109'"
        sql = sql & " and left(dkey,3) = 'DEP' and substring(data,9,len(data)-1) ='" + DACCDEP2.SelectedValue + "'"
        Sql = Sql & " order by unique_id"
        Dim dtReferp As DataTable = uDataBase.GetDataTable(Sql)
        DACCDEP2NO.Text = dtReferp.Rows(0).Item("Data2")

        LMapFile.ImageUrl = DMappath.Text
        LMapFile1.NavigateUrl = DMappath.Text
        LMapFile.Visible = True
        LMapFile1.Visible = True
    End Sub

    Protected Sub DDCOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDCOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DDCOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DCCOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCCOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DCCOST.AutoPostBack = True
            Total()
        End If
    End Sub




    Protected Sub DFCOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFCOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DFCOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DGCOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGCOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DGCOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DHCOST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DHCOST.TextChanged
        If (wStep = 30) Or (wStep = 35) Then
            DHCOST.AutoPostBack = True
            Total()
        End If
    End Sub

    Protected Sub DTYPE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTYPE.SelectedIndexChanged
        'If DTYPE.SelectedValue = "依賴" Then
        '    DACCDEP12.Visible = True
        '    DACCDEP13.Visible = True
        'Else
        '    DACCDEP12.Visible = False
        '    DACCDEP13.Visible = False
        'End If

        Dim SQL As String
        Dim i As Integer
        If DTYPE.SelectedValue <> "依賴" Then
            SQL = "  Select DATA from M_referp"
            SQL = SQL & " where  cat = '3109'"
            SQL = SQL & " and dkey = 'JUDGE'"
            SQL = SQL & " order by unique_id"
        Else
            SQL = " SELECT*  from M_referp where cat ='3109' and DKEY ='RELY' "
        End If
       

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DTYPE1.Items.Clear()
        DTYPE1.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DTYPE1.Items.Add(ListItem1)
        Next
        If DTYPE.SelectedValue <> "依賴" Then
            DTYPE1.Items.Add("待判")
        End If



    End Sub

    Protected Sub DJUDGE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DJUDGE.SelectedIndexChanged
        Dim SQL As String
        Dim i As Integer

        If Mid(DJUDGE.SelectedValue, 1, 1) = "A" Then
            DMANUREASON.Enabled = True
            DMANUREASON1.Enabled = True

            Sql = "  Select DATA from M_referp"
            Sql = Sql & " where  cat = '3109'"
            Sql = Sql & " and dkey = 'MANUREASON'"
            Sql = Sql & " order by unique_id"

            Dim dtReferp As DataTable = uDataBase.GetDataTable(Sql)
            DMANUREASON.Items.Clear()
            DMANUREASON.Items.Add("")

            For i = 0 To dtReferp.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = dtReferp.Rows(i).Item("Data")
                ListItem1.Value = dtReferp.Rows(i).Item("Data")

                DMANUREASON.Items.Add(ListItem1)
            Next

            dtReferp.Clear()


        Else
            DMANUREASON.Enabled = False
            DMANUREASON1.Enabled = False
            DMANUREASON.Items.Clear()
            DMANUREASON.Items.Add("無")
            DMANUREASON1.Items.Clear()
            DMANUREASON1.Items.Add("無")
        End If



    
        SQL = "  SELECT data  FROM M_referp where cat ='3109' and left(dkey,9)='JUDGETYPE' and substring(dkey,11,1) ='" + Mid(DJUDGE.SelectedValue, 1, 1) + "' order by Unique_id "

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DJUDGE1.Items.Clear()
        DJUDGE1.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")

            DJUDGE1.Items.Add(ListItem1)
        Next


    End Sub

   
    Protected Sub DACCDEP12_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DACCDEP12.SelectedIndexChanged
        'ACCDEP13
        Dim SQL As String
        Dim i As Integer
        Sql = "  select  data  from  M_referp where cat ='3109' and left(Dkey,4) ='What'"
        Sql = Sql & " and   left(Dkey,4)='What'"
        SQL = SQL & " and substring(dkey,6,len(dkey)-1) ='" + Trim(DACCDEP12.SelectedValue) + "' order by Unique_id "



        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DACCDEP13.Items.Clear()
        DACCDEP13.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")

            DACCDEP13.Items.Add(ListItem1)
        Next
   

    End Sub

    Protected Sub DJUDGE1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DJUDGE1.SelectedIndexChanged

        Dim Sql As String
        Dim i As Integer

        Sql = "  SELECT  Data  FROM M_referp where cat ='3109' and  substring(dkey,9,len(dkey)-1) ='" + Trim(DJUDGE1.SelectedValue) + "' order by Unique_id "

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(Sql)
        DJUDGE2.Items.Clear()
        DJUDGE2.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")

            DJUDGE2.Items.Add(ListItem1)
        Next

    End Sub


End Class
