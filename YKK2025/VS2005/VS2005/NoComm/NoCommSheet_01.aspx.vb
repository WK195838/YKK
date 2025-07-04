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


Partial Class NoCommSheet_01
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
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String
   

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置


        If Not Me.IsPostBack Then   '不是PostBack
            NewAttachFilePath()

            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件
            '自動起單
            If wStep = 1 Then
                AUTOCREAE()
            End If

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
            Top = 800
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 800
                    End If
                End If
            End If
        Else
            Top = 800
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



        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top & "px"
            DDecideDesc.Style("top") = Top + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top + 200 & "px"
        BNG1.Style("top") = Top + 200 & "px"
        BNG2.Style("top") = Top + 200 & "px"
        BOK.Style("top") = Top + 200 & "px"

        Top = Top + 200
        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 30
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
        Response.Cookies("PGM").Value = "NocommSheet_01.aspx"
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
                Top = 800
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
            Top = 800
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

        Top = Top + 100
        BSAVE.Style("top") = Top + 32 & "px"
        BNG1.Style("top") = Top + 32 & "px"
        BNG2.Style("top") = Top + 32 & "px"
        BOK.Style("top") = Top + 32 & "px"
        DHistoryLabel.Style("top") = Top + 32 + 10 & "px"
        GridView2.Style("top") = Top + 32 + 25 & "px"
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
        Dim Sql As String
        Dim pCode As String
        Dim pCust As String
        Dim pSales As String
        pCode = Trim(Request.QueryString("pItem"))
        pCust = Trim(Request.QueryString("pCust"))
        pSales = Trim(Request.QueryString("pSales"))

        If wStep = 1 Then
            'Sql = " select  sales,name,cust,custname,Item,itemname1,sum(qty)qty,convert(int,sum(amount))amount,'2412021800' as COMMENT"
            'Sql = Sql + "  from  F_nocommlist where sales='" + psales + "' and cust ='" + Request.QueryString("pCust") + "' and item ='" + Request.QueryString("pItem") + "'"
            'Sql = Sql + " group by sales,name,cust,custname,Item,itemname1"



            'Sql = "   select  sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,'' COMMENT,sum(qty)qty,sum(amount)amount "
            'Sql = Sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=0  "
            'Sql = Sql + " and sales='" + pSales + "' and cust ='" + pcust + "' and item ='" + pCode + "'"
            'Sql = Sql + " Group by          sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm"
            'Sql = Sql + " UNION ALL"
            'Sql = Sql + "  select  sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,'' COMMENT,sum(qty)qty,sum(amount)amount "
            'Sql = Sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=1  "

            'If pCode = "" Then
            '    Sql = Sql + " and sales='" + pSales + "' and cust ='" + pCust + "'"
            'Else
            '    Sql = Sql + " and sales='" + pSales + "' and cust ='" + pCust + "' and item ='" + pCode + "'"
            'End If

            '' Sql = Sql + " and sales='" + psales + "' and cust ='" + pcust+ "'"
            'Sql = Sql + " Group by          sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm"
            'Sql = Sql + " UNION ALL"
            'Sql = Sql + " SELECT sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,COMMENT1,sum(qty)qty,sum(amount)amount  FROM ("
            'Sql = Sql + " select  *,SUBSTRING(COMMENT, CHARINDEX( '客訴單號',comment )+5,10)COMMENT1 from F_nocommlist"
            'Sql = Sql + " WHERE  CHARINDEX('客訴單號',comment )>0"
            'Sql = Sql + " and sales='" + pSales + "' and cust ='" + pCust + "' and item ='" + pCode + "'"
            'Sql = Sql + " )A GROUP BY sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,COMMENT1"


            Sql = " select   sales,name,cust,custname,sample,COMMENT,sum(qty)qty,sum(amount)amount from("
            Sql = Sql + "   select   cust,custname,sales,name,confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT','' as COMMENT  "
            Sql = Sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=0  "
            Sql = Sql + "  and cust ='" + pCust + "' and item ='" + pCode + "'"

            Sql = Sql + " UNION ALL"
            Sql = Sql + "  select   cust,custname,sales,name,confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT','' as COMMENT  "
            Sql = Sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=1  "
            If pCode = "" Then
                Sql = Sql + " and cust ='" + pCust + "'"
            Else
                Sql = Sql + " and cust ='" + pCust + "' and item ='" + pCode + "'"

            End If

            Sql = Sql + " UNION ALL"
            Sql = Sql + " SELECT  cust,custname,sales,name,confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT1   FROM ("
            Sql = Sql + " select  *,SUBSTRING(COMMENT, CHARINDEX( '客訴單號',comment )+5,10)COMMENT1 from F_nocommlist"
            Sql = Sql + " WHERE  CHARINDEX('客訴單號',comment )>0"
            Sql = Sql + " and cust ='" + pCust + "' and item ='" + pCode + "'"
            Sql = Sql + " )A  "
            Sql = Sql + " )a GROUP BY sales,name,cust,custname,sample,COMMENT "






            Dim DBUser As DataTable = uDataBase.GetDataTable(Sql)
            '申請者及報告者的資料
            If DBUser.Rows(0).Item("name") = "東莞蘇立傑" Then
                DEmpName.Text = "蘇立傑"
            Else
                DEmpName.Text = DBUser.Rows(0).Item("name")
            End If

            ' DEmpName.Text = "林維珊"
            DCust.Text = DBUser.Rows(0).Item("Cust")
            DCustName.Text = DBUser.Rows(0).Item("CustName")
            '  DCustName.Text = "泰萬億企業有限公司"

            DItem.Text = pCode
            DQty.Text = DBUser.Rows(0).Item("Qty")
            DAmount.Text = DBUser.Rows(0).Item("Amount")

            DComplainNo.Text = DBUser.Rows(0).Item("COMMENT")


            Sql = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
            Sql = Sql & "  where a.empname =b.username"
            Sql = Sql & " and b.userid = '" & Request.QueryString("pApplyID") & "'"
            Sql = Sql & "   And Active = '1' "
            Dim DBUser1 As DataTable = uDataBase.GetDataTable(Sql)

            DAppName.Text = DBUser1.Rows(0).Item("empname")

            Sql = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
            Sql = Sql & "  where a.empname =b.username"
            Sql = Sql & " and b.username = '" & DEmpName.Text & "'"
            Sql = Sql & "   And Active = '1' "
            Dim DBUser2 As DataTable = uDataBase.GetDataTable(Sql)

            DDivision.Text = DBUser2.Rows(0).Item("DepName")

            '細項連結
            LNoCommList.NavigateUrl = "PNoCommList.aspx?pCust=" + pCust + "&pcode=" + pCode

            DCDescrption.Text = ""
            DEDescrption.Text = ""
            '客訴內容
            If DComplainNo.Text <> "" Then
                Sql = " Select * from F_ComplaintOutSheet "
                Sql = Sql & "  where no='" + DComplainNo.Text + "'"

                Dim DBUser3 As DataTable = uDataBase.GetDataTable(Sql)
                If DBUser3.Rows.Count > 0 Then

                    LQCNo.Visible = True
                    LQCNo.Text = "LINK"
                    '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
                    LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(DBUser3.Rows(0).Item("formsno"))

                    DCDescrption.Text = DBUser3.Rows(0).Item("CPCONTENT")
                    DEDescrption.Text = DBUser3.Rows(0).Item("EPCONTENT")
                End If
            End If
            '判斷是否為外注
            If Request.QueryString("pSupplier") = "1" Then
                DSUPPLIER.Checked = True
            End If

        End If




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
            Case 2  '修改
                DNo.Visible = True
                DNo.BackColor = Color.Yellow
                DNo.ReadOnly = False
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = SetNo(wFormSno)

        'Date
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.Visible = True
                DDate.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDate.Visible = True
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = False
            Case 2  '修改
                DDate.Visible = True
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = False
            Case Else   '隱藏
                DDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時



        'Division
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.Visible = True
                DDivision.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select



        'AppName
        Select Case FindFieldInf("AppName")
            Case 0  '顯示
                DAppName.BackColor = Color.LightGray
                DAppName.Visible = True
                DAppName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAppName.BackColor = Color.GreenYellow
                DAppName.Visible = True
            Case 2  '修改
                DAppName.BackColor = Color.Yellow
                DAppName.Visible = True
            Case Else   '隱藏
                DAppName.Visible = False
        End Select


        'EmpName
        Select Case FindFieldInf("EmpName")
            Case 0  '顯示
                DEmpName.BackColor = Color.LightGray
                DEmpName.Visible = True
                DEmpName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEmpName.BackColor = Color.GreenYellow
                DEmpName.Visible = True
            Case 2  '修改
                DEmpName.BackColor = Color.Yellow
                DEmpName.Visible = True
            Case Else   '隱藏
                DEmpName.Visible = False
        End Select



        '"Cust
        Select Case FindFieldInf("Cust")
            Case 0  '顯示
                DCust.BackColor = Color.LightGray
                DCust.Visible = True
                DCust.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCust.BackColor = Color.GreenYellow
                DCust.Visible = True
            Case 2  '修改
                DCust.BackColor = Color.Yellow
                DCust.Visible = True
            Case Else   '隱藏
                DCust.Visible = False
        End Select




        '"CustName
        Select Case FindFieldInf("CustName")
            Case 0  '顯示
                DCustName.BackColor = Color.LightGray
                DCustName.Visible = True
                DCustName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCustName.BackColor = Color.GreenYellow
                DCustName.Visible = True
            Case 2  '修改
                DCustName.BackColor = Color.Yellow
                DCustName.Visible = True
            Case Else   '隱藏
                DCustName.Visible = False
        End Select

        'Qty
        Select Case FindFieldInf("Qty")
            Case 0  '顯示
                DQty.BackColor = Color.LightGray
                DQty.Visible = True
                DQty.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DQty.Visible = True
                DQty.BackColor = Color.GreenYellow
                DQty.ReadOnly = False

                ShowRequiredFieldValidator("DQTYRqd", "DQTY", "異常：請輸入QTY")
            Case 2  '修改
                DQty.Visible = True
                DQty.BackColor = Color.Yellow
                DQty.ReadOnly = False
            Case Else   '隱藏
                DQty.Visible = False
        End Select


        'Amount
        Select Case FindFieldInf("Amount")
            Case 0  '顯示
                DAmount.BackColor = Color.LightGray
                DAmount.Visible = True
                DAmount.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAmount.Visible = True
                DAmount.BackColor = Color.GreenYellow
                DAmount.ReadOnly = False
            Case 2  '修改
                DAmount.Visible = True
                DAmount.BackColor = Color.Yellow
                DAmount.ReadOnly = False
            Case Else   '隱藏
                DAmount.Visible = False
        End Select



        'SReason
        Select Case FindFieldInf("SReason")
            Case 0  '顯示
                DSReason.Visible = True

                DSReason.BackColor = Color.LightGray
                DSReason.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DSReason.BackColor = Color.GreenYellow
                DSReason.Visible = True

                ShowRequiredFieldValidator("DSReasonRqd", "DSReason", "異常：請輸入無償理由")

            Case 2  '修改
                DSReason.Visible = True
                DSReason.BackColor = Color.Yellow

            Case Else   '隱藏
                DSReason.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("SReason", "ZZZZZZ")




        'ComplainNo
        Select Case FindFieldInf("ComplainNo")
            Case 0  '顯示
                DComplainNo.BackColor = Color.LightGray
                DComplainNo.Visible = True
                DComplainNo.Attributes.Add("readonly", "true")
                BQCNO.Visible = False
            Case 1  '修改+檢查
                DComplainNo.Visible = True
                DComplainNo.BackColor = Color.GreenYellow
                DComplainNo.ReadOnly = False
                ShowRequiredFieldValidator("DComplainNoRqd", "DComplainNo", "異常：需輸入客訴單號")
                BQCNO.Visible = True
            Case 2  '修改
                DComplainNo.Visible = True
                DComplainNo.BackColor = Color.Yellow
                DComplainNo.ReadOnly = False
                BQCNO.Visible = True
            Case Else   '隱藏
                DComplainNo.Visible = False
                BQCNO.Visible = False
        End Select


        'CDescrption
        Select Case FindFieldInf("CDescrption")
            Case 0  '顯示
                DCDescrption.BackColor = Color.LightGray
                DCDescrption.Visible = True
                DCDescrption.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DCDescrption.Visible = True
                DCDescrption.BackColor = Color.GreenYellow
                DCDescrption.ReadOnly = False
                ShowRequiredFieldValidator("DCDescrptionRqd", "DCDescrption", "異常：需輸入無償說明")
            Case 2  '修改
                DCDescrption.Visible = True
                DCDescrption.BackColor = Color.Yellow
                DCDescrption.ReadOnly = False
            Case Else   '隱藏
                DCDescrption.Visible = False
        End Select

        'EDescrption
        Select Case FindFieldInf("EDescrption")
            Case 0  '顯示
                DEDescrption.BackColor = Color.LightGray
                DEDescrption.Visible = True
                DEDescrption.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DEDescrption.Visible = True
                DEDescrption.BackColor = Color.GreenYellow
                DEDescrption.ReadOnly = False
                ShowRequiredFieldValidator("DEDescrptionRqd", "DEDescrption", "異常：需輸入無償說明")
            Case 2  '修改
                DEDescrption.Visible = True
                DEDescrption.BackColor = Color.Yellow
                DEDescrption.ReadOnly = False
            Case Else   '隱藏
                DEDescrption.Visible = False
        End Select


        'Comment1
        Select Case FindFieldInf("Comment1")
            Case 0  '顯示
                DComment1.BackColor = Color.LightGray
                DComment1.Visible = True
                DComment1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DComment1.Visible = True
                DComment1.BackColor = Color.GreenYellow
                DComment1.ReadOnly = False
                ShowRequiredFieldValidator("DComment1Rqd", "DComment1", "異常：課長需需輸入說明")
            Case 2  '修改
                DComment1.Visible = True
                DComment1.BackColor = Color.Yellow
                DComment1.ReadOnly = False
            Case Else   '隱藏
                DComment1.Visible = False
        End Select


        'Comment2
        Select Case FindFieldInf("Comment2")
            Case 0  '顯示
                DComment2.BackColor = Color.LightGray
                DComment2.Visible = True
                DComment2.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DComment2.Visible = True
                DComment2.BackColor = Color.GreenYellow
                DComment2.ReadOnly = False
                ShowRequiredFieldValidator("DComment2Rqd", "DComment2", "異常：經理需需輸入說明")
            Case 2  '修改
                DComment2.Visible = True
                DComment2.BackColor = Color.Yellow
                DComment2.ReadOnly = False
            Case Else   '隱藏
                DComment2.Visible = False
        End Select



        'Comment3
        Select Case FindFieldInf("Comment3")
            Case 0  '顯示
                DComment3.BackColor = Color.LightGray
                DComment3.Visible = True
                DComment3.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DComment3.Visible = True
                DComment3.BackColor = Color.GreenYellow
                DComment3.ReadOnly = False
                ShowRequiredFieldValidator("DComment3Rqd", "DComment3", "異常：日籍主管需輸入說明")
            Case 2  '修改
                DComment3.Visible = True
                DComment3.BackColor = Color.Yellow
                DComment3.ReadOnly = False
            Case Else   '隱藏
                DComment3.Visible = False
        End Select


        Select Case FindFieldInf("Attachfile1")
            Case 0  '顯示            

                DAttachfile1.Visible = True



            Case 1  '修改+檢查
                DAttachfile1.Visible = True

            Case 2  '修改
                DAttachfile1.Visible = True


            Case Else   '隱藏
                DAttachfile1.Visible = False

        End Select



        '外注
        Select Case FindFieldInf("SUPPLIER")
            Case 0  '顯示

                DSUPPLIER.Enabled = False

            Case 1  '修改+檢查

                DSUPPLIER.Enabled = True

            Case 2  '修改

                DSUPPLIER.Enabled = True
            Case Else   '隱藏

                DSUPPLIER.Enabled = False

        End Select


    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        'Dim Path As String = "http://localhost:60679/EApproval/Document/001172/"  '測試環境

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http1") & _
        System.Configuration.ConfigurationManager.AppSettings("EApprovalPath")

        Dim SQL As String
        SQL = "Select * From F_NoCommSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DDate.Text = dtData.Rows(0).Item("Date")
            DDivision.Text = dtData.Rows(0).Item("Division")
            DAppName.Text = dtData.Rows(0).Item("AppName")
            DEmpName.Text = dtData.Rows(0).Item("EmpName")
            DCust.Text = dtData.Rows(0).Item("Cust")
            DCustName.Text = dtData.Rows(0).Item("CustName")

            If dtData.Rows(0).Item("Supplier") = "1" Then
                DSUPPLIER.Checked = True
            Else
                DSUPPLIER.Checked = False
            End If

            '細項連結
            LNoCommList.NavigateUrl = "PNoCommList.aspx?pCust=" + DCust.Text + "&pcode=" + dtData.Rows(0).Item("Item") + "&pNo=" + dtData.Rows(0).Item("No")

            DComplainNo.Text = dtData.Rows(0).Item("ComplainNo")
            DQty.Text = dtData.Rows(0).Item("Qty")
            DAmount.Text = dtData.Rows(0).Item("Amount")
            SetFieldData("SReason", dtData.Rows(0).Item("SReason"))
            DCDescrption.Text = dtData.Rows(0).Item("CDescrption")
            DEDescrption.Text = dtData.Rows(0).Item("EDescrption")
            DComment1.Text = dtData.Rows(0).Item("Comment1")
            DComment2.Text = dtData.Rows(0).Item("Comment2")
            DComment3.Text = dtData.Rows(0).Item("Comment3")




            sql = "  select no,formsno,spec,CPCONTENT,EPCONTENT from F_ComplaintOutSheet where no = '" & DComplainNo.Text & "'"


            Dim DBQC As DataTable = uDataBase.GetDataTable(sql)
            If DBQC.Rows.Count > 0 Then
 
                LQCNo.Visible = True
                LQCNo.Text = "LINK"
                '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
                LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(DBQC.Rows(0).Item("formsno"))
            
            End If






            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '1172'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/附件"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "\\" + DBAdapter3.Rows(0).Item("Data1") + DNo.Text + "\附件"

            Dim dirInfo As New System.IO.DirectoryInfo(OpenDir2)

            'Dim FileDir As Integer  '資料夾
            'FileDir = dirInfo.GetDirectories("*").Length
            'Dim FileCount As Integer '檔案
            'FileCount = dirInfo.GetFiles("*.*").Length

            '檢查目錄是否存，不存在就新建一個
            Dim tempFolderPath As String
            tempFolderPath = OpenDir2
            '\\10.245.1.6\wfs$\NoComm\001172\2502250001\附件
            Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
            If dInfo.Exists Then

            Else
                'dInfo.Create()
                My.Computer.FileSystem.CreateDirectory(tempFolderPath)
            End If


            chktemp.Text = tempFolderPath

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
                'DDecideDesc.Text = dtWaitHandle.Rows(0)("DecideDesc")           '說明

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
        Dim Sql As String = ""
        Dim idx As Integer = FindFieldInf(pFieldName)
        Dim i As Integer

        '製造1
        If pFieldName = "SReason" Then
            DSReason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSReason.Items.Add(ListItem1)
                End If
            Else
                Sql = "Select  Data  From M_Referp Where Cat='1172' and dkey ='SReason'  Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(Sql)

                DSReason.Items.Add("")
                For i = 0 To dtReasonCode.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Trim(dtReasonCode.Rows(i)("Data"))
                    ListItem1.Value = Trim(dtReasonCode.Rows(i)("Data"))
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSReason.Items.Add(ListItem1)
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
                Sql = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(Sql)
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
            Sql = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim dtReason As DataTable = uDataBase.GetDataTable(Sql)
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

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("z-index", "137")
        rqdVal.Style.Add("Top", "780px")
        rqdVal.Style.Add("Left", "30px")
        rqdVal.Style.Add("Width", "450px")
        rqdVal.Style.Add("position", "absolute")
        rqdVal.Style.Add("display", "inline")
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
        BQCNO.Attributes.Add("onClick", "GetQCNo()") '取客訴編號

    End Sub
    '##-1
    '##AgentApprovProc-Start
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AgentApprovProc)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub AgentApprovProc(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0
        '------------------------------------------------
        '處理button說明
        'StsDes
        Dim wStsDes As String = ""
        If pSts = "1" Then wStsDes = BOK.Text
        If pSts = "2" Then wStsDes = BNG1.Text
        If pSts = "3" Then wStsDes = BNG2.Text
        '--
        '------------------------------------------------
        '指定下工程簽核者
        'AllocteID
        Dim wAllocateID As String = ""
        '--
        '------------------------------------------------
        '主表單
        'TabelName
        Dim wTableName As String = "F_NoCommSheet"
        '--
        '------------------------------------------------
        '代理簽核
        'RunAgentApprov
        ErrCode = oCommon.RunAgentApprov(pFun, _
                                            pAction, _
                                            pSts, _
                                            wStsDes, _
                                            wFormNo, _
                                            wFormSno, _
                                            wStep, _
                                            wSeqNo, _
                                            wDecideCalendar, _
                                            Request.QueryString("pUserID"), _
                                            wApplyID, _
                                            wAgentID, _
                                            wAllocateID, _
                                            uCommon.ReplaceString(DDecideDesc.Text), _
                                            wTableName, _
                                            wUserIP)
        '--
        '------------------------------------------------
        '表單資料
        If ErrCode = 0 Then
            ModifyData("AGENTAPPROVE", "0")
            AddCommissionNo(wFormNo, wFormSno)
            '
            Dim URL As String = "http://10.245.1.10/WorkFlow/WaitHandle.aspx?pUserID=" + Request.QueryString("pUserID")
            Response.Redirect(URL)
        Else
            EnabledButton()   '起動Button運作
            uJavaScript.PopMsg(Me, "代理簽核異常,請確認或連絡系統人員!")
        End If
    End Sub
    '##AgentApprovProc-End
    '##
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

                '報告者
                If wStep = 1 Then
                    wAllocateID = GetUerID(DEmpName.Text)
                ElseIf wStep = 20 And pFun = "OK" Then  '主管1
                    wAllocateID = GetRelated(GetUerID(DEmpName.Text))
                ElseIf wStep = 30 And pFun = "OK" Then  '主管2
                    wAllocateID = GetRelated(Request.QueryString("pUserID"))

                    If DSUPPLIER.Checked = True Then  '外注主管1後直接結束
                        pAction = 3
                    Else
                        If CInt(DAmount.Text) > 30000 Then
                            pAction = 0
                        Else
                            pAction = 2
                        End If

                    End If
                   
                ElseIf wStep = 40 And pFun = "OK" Then  '日籍
                    If wAgentID <> "" Then
                        wAllocateID = GetRelated(wAgentID)
                    Else
                        wAllocateID = GetRelated(Request.QueryString("pUserID"))
                    End If

                    If CInt(DAmount.Text) > 50000 Then
                        pAction = 0
                    Else
                        pAction = 2
                    End If
                ElseIf wStep = 20 And pFun = "NG1" Then  '退回報告者
                    wAllocateID = GetUerID(DEmpName.Text)
                ElseIf wStep = 30 And pFun = "NG1" Then  '退回報告者
                    wAllocateID = GetUerID(DEmpName.Text)
                ElseIf wStep = 40 And pFun = "NG1" Then  '退回報告者
                    wAllocateID = GetUerID(DEmpName.Text)
                ElseIf wStep = 50 And pFun = "NG1" Then  '退回報告者
                    wAllocateID = GetUerID(DEmpName.Text)
                End If

           
                If wStep = 31 Then
                    '日籍
                    wAllocateID = GetRelated(GetRelated(Request.QueryString("pUserID")))
                End If
            
                If wStep = 50 Then
                    ' 報告者
                    wAllocateID = GetUerID(DEmpName.Text)
                End If









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
            'oCommon.SendMail()

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


        Dim DBDataSet1 As New DataSet

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("TP_Conn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期

        'Dim Path As String = "D:/Yu_Hao Peng Data/Documents/Visual Studio 2005/WebSites/EApproval/Document/001172/" '測試環境
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
        System.Configuration.ConfigurationManager.AppSettings("EApprovalPath")


        Dim sql As String = ""
        sql = " Insert into F_NoCommSheet (Sts, CompletedTime, FormNo, FormSno,"
        sql = sql + " NO,Date,AppName,Division,EmpName,Cust,CustName,Item,ComplainNo,Qty,Amount,"
        sql = sql + "SReason,CDescrption,EDescrption,Comment1,Comment2,Comment3,SUPPLIER,"
        sql = sql + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + "VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日
        sql = sql + " '001172', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號
        sql = sql + "'" + DNo.Text + "',"
        sql = sql + "'" + DDate.Text + "',"
        sql = sql + "'" + Trim(DAppName.Text) + "',"
        sql = sql + "'" + Trim(DDivision.Text) + "',"
        sql = sql + "'" + Trim(DEmpName.Text) + "',"
        sql = sql + "'" + Trim(DCust.Text) + "',"
        sql = sql + "'" + Trim(DCustName.Text) + "',"
        sql = sql + "'" + Trim(DItem.Text) + "',"
        sql = sql + "'" + Trim(DComplainNo.Text) + "',"
        sql = sql + "'" + DQty.Text + "',"
        sql = sql + "'" + DAmount.Text + "',"
        sql = sql + "'" + DSReason.Text + "',"
        sql = sql + " N'" + YKK.ReplaceString(DCDescrption.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DEDescrption.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DComment1.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DComment2.Text) + "', "
        sql = sql + " N'" + YKK.ReplaceString(DComment3.Text) + "', "
        If DSUPPLIER.Checked = True Then
            sql = sql + "1,"
        Else
            sql = sql + "0,"
        End If
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)

        '新增細項
        sql = " delete from  F_NoCommSheetDT"
        sql = sql + " where No='" + DNo.Text + "'"
        uDataBase.ExecuteNonQuery(sql)


        sql = " Insert into F_NoCommSheetDT(Sts, CompletedTime, FormNo, FormSno,"
        sql = sql + " No,Date,ORDDATE,ORDERNO,SAMPLE,BUYER,BUYERNAME,ITEM,ITEMNAME,COLOR,QTY,AMOUNT,COMMENT,"
        sql = sql + " CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + " select '0','" + NowDateTime + "','001172','" + CStr(NewFormSno) + "','" + DNo.Text + "','" + DDate.Text + "',"
        sql = sql + " confirm1,order1,SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 ,COLOR,QTY,amount , COMMENT,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "', "
        sql = sql + "'" & NowDateTime & "' "
        sql = sql + "  from("
        sql = sql + "   select    confirm1  ,order1   ,SAMPLE,BUYER,BUYERNAME,ITEM,itemname1  ,COLOR,QTY,amount as 'AMOUNT',COMMENT   "
        sql = sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=0  "
        sql = sql + "  and cust ='" + DCust.Text + "' and item ='" + DItem.Text + "'"

        sql = sql + " UNION ALL"
        sql = sql + "  select   confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT  "
        sql = sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=1  "
        If DItem.Text = "" Then
            sql = sql + " and cust ='" + DCust.Text + "'"
        Else
            sql = sql + " and cust ='" + DCust.Text + "' and item ='" + DItem.Text + "'"

        End If
        sql = sql + " UNION ALL"
        sql = sql + " SELECT  confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT1   FROM ("
        sql = sql + " select  *,SUBSTRING(COMMENT, CHARINDEX( '客訴單號',comment )+5,10)COMMENT1 from F_nocommlist"
        sql = sql + " WHERE  CHARINDEX('客訴單號',comment )>0"
        sql = sql + " and cust ='" + DCust.Text + "' and item ='" + DItem.Text + "'"
        sql = sql + " )A  "
        sql = sql + " )a"




        uDataBase.ExecuteNonQuery(sql)

     


 
        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '1172'"
        sql = sql + " and dkey ='AttachfilePath' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter1.Rows.Count > 0 Then
            sourceDir = "\\" + DBAdapter1.Rows(0).Item("Data1") + D3.Text   '來源
        End If

        '主檔資料
        sql = " select  data,replace(data,'/','\')data1  from M_referp"
        sql = sql + " where cat = '1172'"
        sql = sql + " and dkey ='AttachfilePath1' "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sql)

        If DBAdapter2.Rows.Count > 0 Then
            backupDir = "\\" + DBAdapter2.Rows(0).Item("Data1") + DNo.Text      '目的  
        End If

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


        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                           CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        'Dim Path As String = "D:/Yu_Hao Peng Data/Documents/Visual Studio 2005/WebSites/EApproval/Document/001172/" '測試環境

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
        System.Configuration.ConfigurationManager.AppSettings("EApprovalPath")

        Dim sql As String


        sql = " Update F_NoCommSheet"
        sql = sql + " Set "
        '##-2
        '##AgentApprov-Start
        'If pFun <> "SAVE" Then
        '    sql = sql + " Sts = '" & pSts & "',"
        '    sql = sql + " CompletedTime = '" & NowDateTime & "',"
        'End If
        '
        If pFun <> "SAVE" And pFun <> "AGENTAPPROVE" Then
            sql = sql + " Sts = '" & pSts & "',"
            sql = sql + " CompletedTime = '" & NowDateTime & "',"
        End If
        '##AgentApprov-End
        '##
        sql = sql + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        sql = sql + " Date = N'" & YKK.ReplaceString(DDate.Text) & "',"
        sql = sql + " AppName = N'" & YKK.ReplaceString(DAppName.Text) & "',"
        sql = sql + " EmpName = N'" & YKK.ReplaceString(DEmpName.Text) & "',"
        sql = sql + " Division = N'" & YKK.ReplaceString(DDivision.Text) & "',"
        sql = sql + " Cust = N'" & YKK.ReplaceString(DCust.Text) & "',"
        sql = sql + " CustName = N'" & YKK.ReplaceString(DCustName.Text) & "',"

        sql = sql + " ComplainNo = N'" & YKK.ReplaceString(DComplainNo.Text) & "',"
        sql = sql + " Qty = N'" & YKK.ReplaceString(DQty.Text) & "',"
        sql = sql + " Amount = N'" & YKK.ReplaceString(DAmount.Text) & "',"
        sql = sql + " SReason= N'" & YKK.ReplaceString(DSReason.SelectedValue) & "',"
        sql = sql + " CDescrption = N'" & YKK.ReplaceString(DCDescrption.Text) & "',"
        sql = sql + " EDescrption= N'" & YKK.ReplaceString(DEDescrption.Text) & "',"
        sql = sql + " Comment1= N'" & YKK.ReplaceString(DComment1.Text) & "',"
        sql = sql + " Comment2= N'" & YKK.ReplaceString(DComment2.Text) & "',"
        sql = sql + " Comment3= N'" & YKK.ReplaceString(DComment3.Text) & "',"
        If DSUPPLIER.Checked = True Then
            sql = sql + "SUPPLIER=1,"
        Else
            sql = sql + "SUPPLIER=0,"
        End If
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
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
            '##-3
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "NG2") = 0 Then
                AgentApprovProc("NG2", 2, "3")
            Else
                FlowControl("NG2", 2, "3")
            End If
            '##AgentApprov-End
            '##
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

        DisabledButton()   '停止Button運作
        '##-4
        '##AgentApprov-Start
        If oCommon.UseAgentApprov(wFormNo, wStep, "NG1") = 0 Then
            AgentApprovProc("NG1", 1, "2")
        Else
            FlowControl("NG1", 1, "2")
        End If
        '##AgentApprov-End
        '##
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
            '##-5
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "OK") = 0 Then
                AgentApprovProc("OK", 0, "1")
            Else
                FlowControl("OK", 0, "1")
            End If
            '##AgentApprov-End
            '##
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub
    '自動起單
    Sub AUTOCREAE()
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            '##-5
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "OK") = 0 Then
                AgentApprovProc("OK", 0, "1")
            Else
                FlowControl("OK", 0, "1")
            End If
            '##AgentApprov-End
            '##
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
        Dim allowedExtensions As String() = {".pdf", ".xls", ".xlsx"}   '定義允許的檔案格式
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
            If UPFile.PostedFile.ContentLength <= 3000 * 1024 Then
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
                ErrCode = oCommon.CommissionNo("001172", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If


        '檢查資料夾是否有檔案

        'If wStep = 10 Or wStep = 500 Then

        '    If ErrCode = 0 Then
        '        Dim dirInfo As New System.IO.DirectoryInfo(chktemp.Text)
        '        Dim FileDir As Integer  '資料夾
        '        FileDir = dirInfo.GetDirectories("*").Length
        '        Dim FileCount As Integer '檔案
        '        FileCount = dirInfo.GetFiles("*.*").Length
        '        If FileCount > 0 Or FileDir > 0 Then
        '        Else
        '            ErrCode = 9065
        '        End If
        '    End If

        'End If


        Try

        Catch ex As Exception
            ErrCode = 9075
        End Try

        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "Item資料有誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!(限PDF EXCEL)"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!(限5MB)"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"
            If ErrCode = 9070 Then Message = "發現有空白重覆,請重新檢測資料,請確認!"
            If ErrCode = 9071 Then Message = "Item Name(2)字串過長(>34),請重新檢測資料,請確認!"
            If ErrCode = 9072 Then Message = "發現特殊要求中有未排序資料,請重新檢測資料,請確認!"
            If ErrCode = 9073 Then Message = "溫度不正常，請確認!"
            If ErrCode = 9074 Then Message = "溼度不正常，請確認!"
            If ErrCode = 9075 Then Message = "非數字格式,請確認!"
            If ErrCode = 9065 Then Message = "請開啟資料夾放入附件!"
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '

        Return isOK

    End Function


    '取得關係人
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID,RRUserID,USERID  from M_Related where userid='" & userId & "' and RelatedID='A'"

        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            '主管1
            NextGate = dt.Rows(0)("RUserID")
        End If
        Return NextGate
    End Function

    '取得關係人
    Function GetUerID(ByVal Username As String) As String

        Dim sql As String

        sql = "select UserID from M_users where username=N'" & Username + "'"


        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then

            NextGate = dt.Rows(0)("UserID")


        End If
        Return NextGate
    End Function

    Protected Sub DQC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DQC.TextChanged

        Dim sql As String
        'sql = "  select no,formsno,spec,CPCONTENT,EPCONTENT from F_ComplaintOutSheet where no = '" & DQC.Text & "'"
        'sql = sql + " and spec in ("
        'sql = sql + " select item from f_nocommsheetdt "
        'sql = sql + " where no = '" + DNo.Text + "' )"

        sql = " select * from F_ComplaintOutSheet where  no = '" & DQC.Text & "'"


        Dim DBQC As DataTable = uDataBase.GetDataTable(sql)
        If DBQC.Rows.Count > 0 Then
            DComplainNo.Text = DBQC.Rows(0).Item("NO")
            DCDescrption.Text = DBQC.Rows(0).Item("CPCONTENT")
            DEDescrption.Text = DBQC.Rows(0).Item("EPCONTENT")
            LQCNo.Visible = True
            LQCNo.Text = "LINK"
            '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
            LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(DBQC.Rows(0).Item("formsno"))
        Else
            LQCNo.Visible = False
            DQC.Text = ""
            DComplainNo.Text = ""
            ' uJavaScript.PopMsg(Me, "客訴單號ITEM 與NO COMMLIST ITEM 不符合，請再確認！")
        End If



    End Sub

    Protected Sub DComplainNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DComplainNo.TextChanged


        Dim sql As String
        'sql = "  select no,formsno,spec,CPCONTENT,EPCONTENT from F_ComplaintOutSheet where no = '" & DComplainNo.Text & "'"
        'sql = sql + " and spec in ("
        'sql = sql + " select item from f_nocommsheetdt "
        'sql = sql + " where no = '" + DNo.Text + "' )"

        sql = " select * from F_ComplaintOutSheet where  no = '" & DQC.Text & "'"


        Dim DBQC As DataTable = uDataBase.GetDataTable(sql)
        If DBQC.Rows.Count > 0 Then
            DComplainNo.Text = DBQC.Rows(0).Item("NO")
            DCDescrption.Text = DBQC.Rows(0).Item("CPCONTENT")
            DEDescrption.Text = DBQC.Rows(0).Item("EPCONTENT")
            LQCNo.Visible = True
            LQCNo.Text = "LINK"
            '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
            LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(DBQC.Rows(0).Item("formsno"))
        Else
            LQCNo.Visible = False
            DQC.Text = ""
            DComplainNo.Text = ""
            ' uJavaScript.PopMsg(Me, "客訴單號ITEM 與NO COMMLIST ITEM 不符合，請再確認！")
        End If


    End Sub



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
        SQL = SQL + " where cat = '1172'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        D3.Text = Now.ToString("yyyyMMddHHmmss")




        OpenDir1 = OpenDir1 + D3.Text + "/附件"   '開啟附檔資料夾路徑


        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

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


End Class


