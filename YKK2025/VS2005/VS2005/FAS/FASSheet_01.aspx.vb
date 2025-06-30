Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class FASSheet_01
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
    Dim Count As Integer
    Dim intTotal As Decimal = 0
    Dim AMT As Decimal = 0
    Dim INVQTY As Decimal = 0
    Dim INVAMT As Decimal = 0

    Dim SCount As Decimal = 0
    Dim DetailSQL As String
    ' 使用者ID


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not IsPostBack Then
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
                ShowFormData()          ' 顯示表單資料
                UpdateTranFile()        ' 更新交易資料
            End If
        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查


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
        Top = 500
    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()

        If GridView3.Rows.Count > 0 Then
            Dim GVTop As Integer = GridView3.Style("top").Replace("px", "")
            Dim GVHeight As Integer = GridView3.Rows.Count * 105   ' 55是列高

            Dim ControlTop As Integer = (GVTop + GVHeight)
            If DDescSheet.Visible Then                                      ' 說明
                DDescSheet.Style("top") = ControlTop & "px"
                DDecideDesc.Style("top") = ControlTop + 6 & "px"
                ControlTop += 81
            End If

            If DDelay.Visible Then                                          ' 延遲說明
                DDelay.Style("top") = ControlTop & "px"
                DReasonCode.Style("top") = ControlTop + 9 & "px"
                DReason.Style("top") = ControlTop + 6 & "px"
                DReasonDesc.Style("top") = ControlTop + 39 & "px"
                ControlTop += 161
            End If

            BSAVE.Style("top") = ControlTop & "px"
            BNG1.Style("top") = ControlTop & "px"
            BNG2.Style("top") = ControlTop & "px"
            BOK.Style("top") = ControlTop & "px"

            'ControlTop += 10
            ' GridView2.Visible = False
            If GridView2.Rows.Count > 0 Then                                ' 核定履歷
                DHistoryLabel.Style("top") = ControlTop + 20 & "px"
                ControlTop += 50
                GridView2.Style("top") = ControlTop & "px"
            End If
        ElseIf GridView1.Rows.Count > 0 Then
            Dim GVTop As Integer = GridView1.Style("top").Replace("px", "")
            Dim GVHeight As Integer = GridView1.Rows.Count * 105    ' 55是列高

            Dim ControlTop As Integer = (GVTop + GVHeight)
            If DDescSheet.Visible Then                                      ' 說明
                DDescSheet.Style("top") = ControlTop & "px"
                DDecideDesc.Style("top") = ControlTop + 6 & "px"
                ControlTop += 81
            End If

            If DDelay.Visible Then                                          ' 延遲說明
                DDelay.Style("top") = ControlTop & "px"
                DReasonCode.Style("top") = ControlTop + 9 & "px"
                DReason.Style("top") = ControlTop + 6 & "px"
                DReasonDesc.Style("top") = ControlTop + 39 & "px"
                ControlTop += 161
            End If

            If ControlTop < 300 Then
                ControlTop = ControlTop + 81
            End If
            BSAVE.Style("top") = ControlTop & "px"
            BNG1.Style("top") = ControlTop & "px"
            BNG2.Style("top") = ControlTop & "px"
            BOK.Style("top") = ControlTop & "px"

            'ControlTop += 10
            ' GridView2.Visible = False
            If GridView2.Rows.Count > 0 Then                                ' 核定履歷
                DHistoryLabel.Style("top") = ControlTop + 20 & "px"
                ControlTop += 50
                GridView2.Style("top") = ControlTop & "px"
            End If
        Else
            Dim ControlTop As Integer
            ControlTop = 500
            BSAVE.Style("top") = ControlTop & "px"
            BNG1.Style("top") = ControlTop & "px"
            BNG2.Style("top") = ControlTop & "px"
            BOK.Style("top") = ControlTop & "px"
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

        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        Response.Cookies("PGM").Value = "FASSheet_01.aspx"
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
                BSAVE.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSAVE.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                BNG1.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG1.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
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
        If Not IsPostBack Then
            ' LAttachfile1.Visible = False
        End If
        'BSAVE.Style("top") = Top & "px"
        'BNG1.Style("top") = Top & "px"
        ' BNG2.Style("top") = Top & "px"
        '  BOK.Style("top") = Top & "px"
        '  DHistoryLabel.Style("top") = Top + 32 & "px"
        '  GridView2.Style("top") = Top + 32 + 16 & "px"
       
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

        If pPost = "New" Then
            Dim SQL As String
            DAppDate.Text = CDate(NowDateTime).ToString("yyyy/MM/dd")
            SQL = "Select e.Com_Code,e.Com_Name,e.Dep_Code,e.Dep_Name,e.[ID],e.[Name],e.Job_Title_Code,e.Job_Title from M_Users u inner join M_Emp e on u.EmpID = e.[ID] and u.DepoID = e.Com_Code where u.UserID='" & Request.QueryString("pUserID") & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            '取得申請者資訊
            If dt.Rows.Count > 0 Then
                DDivision.Text = dt.Rows(0)("Dep_Name").ToString.Trim
                DApper.Text = dt.Rows(0)("Name").ToString.Trim
            End If

        End If

        'If wStep = 1 Then
        '細項內容
        'LFCDataList.NavigateUrl = "FCDataList.aspx?Userid=" & wApplyID
        'Else
      
        'End If


        ''細項DETAIL 
        'DetailSQL = " select  FIX,CustomerCode,CustomerName,BuyerCode,BuyerName,REQDATE,KEEPCODE,BUYMONTH,CORDERNO"
        'DetailSQL = DetailSQL + " ,ITEMCODE1,ITEMNAME1,COLOR1,ORDERQTYP1	"
        'DetailSQL = DetailSQL + ",ITEMCODE2,ITEMNAME2,COLOR2,ORDERQTYP2"
        'DetailSQL = DetailSQL + " from  f_fcdata"
        'DetailSQL = DetailSQL + " where createuser  ='" + wApplyID + "'"
        'DetailSQL = DetailSQL + " and convert(date,buymonth+'/1') >convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
        'DetailSQL = DetailSQL + " and formsno =0"
        'DetailSQL = DetailSQL + " and sts =0"
        'DetailSQL = DetailSQL + " order by CustomerCode, BuyerCode,REQDATE,BUYMONTH"

        '細項加總起單時
        If wStep = 1 Then
            Dim SQL1 As String
            SQL1 = " select   customercode+'-'+buyercode as BuyerCode,'一批' as ITEMCODE,sum((convert(int,orderqtyp1))) as qty, sum((convert(int,orderqtyp1*UnitPrice))) as AMT,sum(INVQTY)INVQTY, convert(int,sum(INVAMT)) INVAMT"
            SQL1 = SQL1 + " from  f_fcdata"
            SQL1 = SQL1 + " where createuser  ='" + Request.QueryString("pUserID") + "'"
            SQL1 = SQL1 + " and convert(date,buymonth+'/1') >=convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
            SQL1 = SQL1 + " and formsno =0"
            SQL1 = SQL1 + " and scheck not in ('C','P')"
            SQL1 = SQL1 + " group by CustomerCode, BuyerCode"

            Dim dt1 As DataTable = uDataBase.GetDataTable(SQL1)
            If dt1.Rows.Count > 0 Then
                Count = dt1.Rows.Count
                GridView1.DataSource = dt1
                GridView1.DataBind()
                '細項內容
                LFCDataList.NavigateUrl = "FCDataList.aspx?Userid=" & Request.QueryString("pUserID") & "&pFormSno=" & CStr(wFormSno)
              
            End If
        Else
            '細項內容
            LFCDataList.NavigateUrl = "FCDataList.aspx?Userid=" & Request.QueryString("pUserID") & "&pFormSno=" & CStr(wFormSno)
        End If

    
        SetControlPosition()    ' 設定控制項位置

        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.White
                DNo.Visible = True
                DNo.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNo.Visible = True
                DNo.BackColor = Color.GreenYellow
                DNo.ReadOnly = False
                ShowRequiredFieldValidator("DNoRqd", "DNo", "異常：需輸入Ｎｏ")
            Case 2  '修改
                DNo.Visible = True
                DNo.BackColor = Color.Yellow
                DNo.ReadOnly = False
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""

        'Date
        Select Case FindFieldInf("AppDate")
            Case 0  '顯示
                DAppDate.BackColor = Color.LightGray
                DAppDate.Visible = True
                DAppDate.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAppDate.Visible = True
                DAppDate.BackColor = Color.GreenYellow
                DAppDate.ReadOnly = False
                ShowRequiredFieldValidator("DAppDateRqd", "DAppDate", "異常：需輸入申請日期")
            Case 2  '修改
                DAppDate.Visible = True
                DAppDate.BackColor = Color.Yellow
                DAppDate.ReadOnly = False
            Case Else   '隱藏
                DAppDate.Visible = False
        End Select
        If pPost = "New" Then DAppDate.Text = Now.ToString("yyyy/MM/dd") '現在日時


        'Name
        Select Case FindFieldInf("Apper")
            Case 0  '顯示
                DApper.BackColor = Color.LightGray
                DApper.Visible = True
                DApper.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DApper.Visible = True
                DApper.BackColor = Color.GreenYellow
                DApper.ReadOnly = False
                ShowRequiredFieldValidator("DApperRqd", "DApper", "異常：需輸入姓名")
            Case 2  '修改
                DApper.Visible = True
                DApper.BackColor = Color.Yellow
                DApper.ReadOnly = False
            Case Else   '隱藏
                DApper.Visible = False
        End Select

        'Division
        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.Visible = True
                DDivision.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDivision.Visible = True
                DDivision.BackColor = Color.GreenYellow
                DDivision.ReadOnly = False
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
            Case 2  '修改
                DDivision.Visible = True
                DDivision.BackColor = Color.Yellow
                DDivision.ReadOnly = False
            Case Else   '隱藏
                DDivision.Visible = False
        End Select

        'Remark
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.Visible = True
                DRemark.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRemark.Visible = True
                DRemark.BackColor = Color.GreenYellow
                DRemark.ReadOnly = False
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需輸入申請備註")
            Case 2  '修改
                DRemark.Visible = True
                DRemark.BackColor = Color.Yellow
                DRemark.ReadOnly = False
            Case Else   '隱藏
                DRemark.Visible = False
        End Select

        'Type

        Select Case FindFieldInf("TYPE")
            Case 0  '顯示
                DTYPE.BackColor = Color.LightGray
                DTYPE.Visible = True

            Case 1  '修改+檢查
                DTYPE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTYPERqd", "DTYPE", "異常：需輸入申請理由")
                DTYPE.Visible = True
            Case 2  '修改
                DTYPE.BackColor = Color.Yellow
                DTYPE.Visible = True
            Case Else   '隱藏
                DTYPE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("TYPE", "ZZZZZZ")

    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("FASFilePath")
        Dim SQL As String
        SQL = "Select * From F_FASSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dt.Rows(0).Item("No")
            DApper.Text = dt.Rows(0).Item("Apper")
            DAppDate.Text = dt.Rows(0).Item("AppDate")
            DDivision.Text = dt.Rows(0).Item("Division")
            DRemark.Text = dt.Rows(0).Item("Remark")
            SetFieldData("TYPE", dt.Rows(0).Item("Type"))
        End If

        If wStep > 1 And wStep <> 500 Then
            SQL = "Select Buyer,ITEMCODE,qty,convert(int,UnitPrice)AMT,INVQTY,convert(int,INVAMT)INVAMT,attachfile From F_FASSheetDT "
            SQL &= " Where FormNo =  '" & wFormNo & "'"
            SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
            Dim dtDttail As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                Count = dtDttail.Rows.Count
                GridView3.DataSource = dtDttail
                GridView3.DataBind()
            End If

        Else
            Dim SQL1 As String
            SQL1 = " select   customercode+'-'+buyercode as BuyerCode,'一批' as ITEMCODE,sum((convert(int,orderqtyp1))) as qty,sum((convert(int,orderqtyp1*UnitPrice))) as Amt,sum(INVQTY)INVQTY, convert(int,sum(INVAMT)) INVAMT"
            SQL1 = SQL1 + " from  f_fcdata"
            SQL1 = SQL1 + " where createuser  ='" + Request.QueryString("pUserID") + "'"
            SQL1 = SQL1 + " and convert(date,buymonth+'/1') >=convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
            SQL1 = SQL1 + " and scheck not in ('C','P')"
            SQL1 = SQL1 + " and formsno = '" & CStr(wFormSno) & "'"
            SQL1 = SQL1 + " group by CustomerCode, BuyerCode"
            Dim dt1 As DataTable = uDataBase.GetDataTable(SQL1)
            If dt1.Rows.Count > 0 Then
                Count = dt1.Rows.Count
                GridView1.DataSource = dt1
                GridView1.DataBind()
                '細項內容
                LFCDataList.NavigateUrl = "FCDataList.aspx?Userid=" & Request.QueryString("pUserID") & "&pFormSno=" & CStr(wFormSno)

            End If

        End If


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
        SQL = SQL + "Order by Unique_ID Desc "

        GridView2.DataSource = uDataBase.GetDataTable(SQL)
        GridView2.DataBind()

        SetControlPosition()    ' 設定控制項位置

    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer


        idx = FindFieldInf(pFieldName)
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
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '8001'"
                SQL = SQL & " and dkey = 'TYPE' "
                SQL = SQL & " order by createtime desc "

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DTYPE.Items.Add("")
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

    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        '
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
                        DNo.Text = SetNo(NewFormSno)
                    End If
                    pRunNextStep = 1
                Else
                    If RepeatRun = False Then   '不是通知的重覆執行
                        '更新交易資料
                        ModifyTranData(pFun, pSts)
                        '流程資料結束(表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽))
                        RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)

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
                '= DSPDPerson.SelectedValue

                '取得簽核者
                '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
                RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                              pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)

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
            If wbFormSno = 0 And wbStep < 3 Then    '判斷是否起單
              
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)

            Else
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
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
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("FASFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String = ""
        sql = "INSERT INTO F_FASSheet ( " & _
              "Sts, CompletedTime, FormNo, FormSno, " & _
              "No, AppDate, Apper,Division,Type,Remark, " & _
              "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
                "VALUES("
        sql &= "'0' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & wFormNo & "' ,"
        sql &= "'" & CStr(NewFormSno) & "' ,"
        sql &= "'" & DNo.Text & "' ,"
        sql &= "'" & DAppDate.Text & "' ,"
        sql &= "'" & DApper.Text & "' ,"
        sql &= "'" & DDivision.Text & "' ,"
        sql &= "'" & DTYPE.SelectedValue & "' ,"
        sql &= "'" & DRemark.Text & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' ,"
        sql &= "'" & Request.QueryString("pUserID") & "' ,"
        sql &= "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)


        '先刪除再新增細項
        sql = " delete from  F_FASSheetdt "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "  And FormSno =  '" & CStr(NewFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)

        '新增細項
        For Each di As GridViewRow In GridView1.Rows
         
            Dim File As FileUpload = CType(di.FindControl("FileUpLoad1"), FileUpload)

       
            '細項
            sql = "INSERT INTO F_FASSheetdt ( " & _
                  "FormNo, FormSno, " & _
                  "BUYER,ItemCode,Qty,UnitPrice,INVQTY,INVAMT,Attachfile," & _
                  "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
                    "VALUES("
            sql &= "'" & wFormNo & "' ,"
            sql &= "'" & CStr(NewFormSno) & "' ,"
            sql &= "'" & di.Cells(0).Text & "' ,"
            sql &= "'" & di.Cells(1).Text & "' ,"
            sql &= "'" & CStr(CInt(di.Cells(2).Text)) & "' ,"
            sql &= "'" & CStr(CInt(di.Cells(3).Text)) & "' ,"
            sql &= "'" & CStr(CInt(di.Cells(4).Text)) & "' ,"
            sql &= "'" & CStr(CInt(di.Cells(5).Text)) & "' ,"
            If File.HasFile Then
                sql &= "'" & CStr(NewFormSno) + "-AttachFile-" + File.FileName & "' ,"
                File.SaveAs(Path & CStr(NewFormSno) + "-AttachFile-" + File.FileName)
            Else
                sql &= "'',"
            End If
            sql &= "'" & Request.QueryString("pUserID") & "' ,"
            sql &= "'" & NowDateTime & "' ,"
            sql &= "'" & Request.QueryString("pUserID") & "' ,"
            sql &= "'" & NowDateTime & "' )"
            uDataBase.ExecuteNonQuery(sql)


        Next


        '回寫單號至備料投入
        sql = " update  f_fcdata "
        sql &= " set formno ='" & wFormNo & "'"
        sql &= " ,formsno ='" & CStr(NewFormSno) & "'"
        sql &= " where createuser  ='" + Request.QueryString("pUserID") + "'"
        sql &= " and convert(date,buymonth+'/1') >=convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
        sql &= " and formsno =0"

        uDataBase.ExecuteNonQuery(sql)



    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("FASFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim sql As String = ""
        sql = "Update  F_FASSheet "
        If pFun <> "SAVE" Then
            sql &= " set  Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        sql &= " TYPE='" & DTYPE.Text & "',"
        sql &= " Remark='" & DRemark.Text & "',"
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "'"
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "  And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)

        If wStep = 500 Then

            '先刪除再新增細項
            sql = " delete from  F_FASSheetdt "
            sql &= " Where FormNo  =  '" & wFormNo & "'"
            sql &= "  And FormSno =  '" & CStr(wFormSno) & "'"
            uDataBase.ExecuteNonQuery(sql)

            For Each di As GridViewRow In GridView1.Rows
                Dim BUYER As String
                BUYER = di.Cells(1).Text
                Dim File As FileUpload = CType(di.FindControl("FileUpLoad1"), FileUpload)

                ' If File.HasFile Then
                ' ' File.SaveAs(System.IO.Path.Combine(Server.MapPath("Uploads"), File.FileName))
                ' File.SaveAs("C:\Temp\" & File.FileName)
                'End If

                '細項
                sql = "INSERT INTO F_FASSheetdt ( " & _
                      "FormNo, FormSno, " & _
                      "BUYER,ItemCode,Qty,UnitPrice,INVQTY,INVAMT,Attachfile," & _
                      "CreateUser, CreateTime, ModifyUser, ModifyTime) " & _
                        "VALUES("
                sql &= "'" & wFormNo & "' ,"
                sql &= "'" & CStr(wFormSno) & "' ,"
                sql &= "'" & di.Cells(0).Text & "' ,"
                sql &= "'" & di.Cells(1).Text & "' ,"
                sql &= "'" & CStr(CInt(di.Cells(2).Text)) & "' ,"
                sql &= "'" & CStr(CInt(di.Cells(3).Text)) & "' ,"
                sql &= "'" & CStr(CInt(di.Cells(4).Text)) & "' ,"
                sql &= "'" & CStr(CInt(di.Cells(5).Text)) & "' ,"
                If File.HasFile Then
                    sql &= "'" & File.FileName & "' ,"
                    File.SaveAs(Path & File.FileName)
                Else
                    sql &= "'',"
                End If
                sql &= "'" & Request.QueryString("pUserID") & "' ,"
                sql &= "'" & NowDateTime & "' ,"
                sql &= "'" & Request.QueryString("pUserID") & "' ,"
                sql &= "'" & NowDateTime & "' )"
                uDataBase.ExecuteNonQuery(sql)
            Next
       
        End If
      

        '回寫單號至備料投入
        sql = " update  f_fcdata "
        sql &= " set formno ='" & wFormNo & "'"
        sql &= " ,formsno ='" & CStr(wFormSno) & "'"
        sql &= " where createuser  ='" + Request.QueryString("pApplyID") + "'"
        sql &= " and convert(date,buymonth+'/1') >=convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
        sql &= " and formsno ='" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(sql)

        If (wStep = 20 And pFun = "OK") Or (wStep = 500 And pFun = "NG") Then  '最後一關存履歷
            '  sql = " delete from f_fcdatahistory "
            '  sql &= " where  formsno ='" & CStr(wFormSno) & "'"
            ' uDataBase.ExecuteNonQuery(sql)
            '將單號清空


            sql = " insert into   f_fcdatahistory "
            sql &= " select * from  f_fcdata"
            sql &= " where createuser  ='" + Request.QueryString("pApplyID") + "'"
            sql &= " and convert(date,buymonth+'/1') >=convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
            sql &= " and formsno ='" & CStr(wFormSno) & "'"
            uDataBase.ExecuteNonQuery(sql)

            sql = " update  f_fcdata "
            sql &= " set formno =''"
            sql &= " ,formsno ='' "
            sql &= " ,INVQTY = 0 "
            sql &= " ,INVAMT=0"
            sql &= " ,Q_ScheProd=0"
            sql &= " ,Q_OnProd=0"
            sql &= " ,Q_FreeInv=0"
            sql &= " ,Q_KeepInv=0"
            sql &= " ,UnitPrice=0"
            sql &= " where  formsno ='" & CStr(wFormSno) & "'"
            uDataBase.ExecuteNonQuery(sql)

        End If
    

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
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> dtCommissionNo.Rows(0)("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
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
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
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
        If InputDataOK(3) Then
            DisabledButton()   '停止Button運作
            ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
            ModifyTranData("SAVE", "0")       '更新交易資料
            Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG2.Click
        If InputDataOK(2) Then
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG1.Click
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        If InputDataOK(0) Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     判斷是否可繼續執行(驗證資料)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""

        If pAction = 0 Then
            'Check上傳附件Size及格式
            For Each di As GridViewRow In GridView1.Rows

                Dim File As FileUpload = CType(di.FindControl("FileUpLoad1"), FileUpload)

                If DTYPE.SelectedValue = "客戶, BUYER需求" Then
                    If File.HasFile Then
                        If File.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(File)
                        End If
                    Else
                        ErrCode = "9010"
                    End If
                End If

            Next

            If DTYPE.SelectedValue = "其他" Then
                If DRemark.Text = "" Then
                    ErrCode = "9040"
                End If
            End If

            If GridView1.Rows.Count = 0 And GridView3.Rows.Count = 0 Then
                ErrCode = "9050"
            End If
        End If

       


        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "請上傳附檔!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "請輸入備註!"
            If ErrCode = 9050 Then Message = "無細項不可申請!"
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
        Dim YStr As String = ""
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim i As Integer

        YStr = Mid(CStr(DateTime.Now.Year), 3, 2)  '年    


        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = YStr + Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 4 - 1
            Str1 = "0" + Str1
        Next

        SetNo = "LS" + Str + Str1
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
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".xlsx", ".doc", ".docx", ".ppt"}   '定義允許的檔案格式
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
   
    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("FASFilePath")
        Dim formatNu As Integer

        '附檔連結
        If e.Row.RowType <> DataControlRowType.Footer And e.Row.RowType <> DataControlRowType.Header Then
            If Trim(e.Row.Cells(6).Text) <> "&nbsp;" Then '不是空白
                Dim h1 As New HyperLink
                ' h1.Text = e.Row.Cells(3).Text
                h1.Text = "附檔"
                h1.Target = "_blank"
                h1.NavigateUrl = Path & e.Row.Cells(6).Text  '上傳檔案
                ' e.Row.Cells(3).Text = ""
                e.Row.Cells(6).Controls.Add(h1)
            End If
            intTotal += e.Row.Cells(2).Text
            AMT += e.Row.Cells(3).Text
            INVQTY += e.Row.Cells(4).Text
            INVAMT += e.Row.Cells(5).Text

            SCount = SCount + 1

   
        End If

        e.Row.Cells(7).Visible = False


        If e.Row.RowType <> DataControlRowType.Header Then
            e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(5).HorizontalAlign = HorizontalAlign.Right
        End If



        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(1).Text = FormatNumber(SCount, formatNu, TriState.True, TriState.False, TriState.True) + "件"

            e.Row.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(3).Text = FormatNumber(AMT, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(4).Text = FormatNumber(INVQTY, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(5).Text = FormatNumber(INVAMT, formatNu, TriState.True, TriState.False, TriState.True)

            e.Row.Cells(0).Text = "合計"

        End If
       

    End Sub

 
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound


        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("FASFilePath")
        Dim formatNu As Integer


        If e.Row.RowType <> DataControlRowType.Footer And e.Row.RowType <> DataControlRowType.Header Then
            intTotal += e.Row.Cells(2).Text
            AMT += e.Row.Cells(3).Text
            INVQTY += e.Row.Cells(4).Text
            INVAMT += e.Row.Cells(5).Text

            SCount = SCount + 1
           
        End If

        If e.Row.RowType <> DataControlRowType.Header Then
            e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(5).HorizontalAlign = HorizontalAlign.Right
        End If





        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(1).Text = FormatNumber(SCount, formatNu, TriState.True, TriState.False, TriState.True) + "件"

            e.Row.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(3).Text = FormatNumber(AMT, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(4).Text = FormatNumber(INVQTY, formatNu, TriState.True, TriState.False, TriState.True)
            e.Row.Cells(5).Text = FormatNumber(INVAMT, formatNu, TriState.True, TriState.False, TriState.True)

            e.Row.Cells(0).Text = "合計"

        End If



    End Sub

    
End Class
