Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic


Partial Class DiscountSheet_02
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()          '設定共用參數

        ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查

        ShowFormData()      '顯示表單資料


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
        Response.Cookies("PGM").Value = "DiscountSheet_01.aspx"
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
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性


    End Sub
   
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()


        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                             System.Configuration.ConfigurationManager.AppSettings("DiscountPath")  'WIS-TempPath

        Dim SQL As String
        SQL = "Select * From F_DiscountSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No

            '開啟報廢資料檔

            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3102'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text
            End If

            'Dim OpenDir As String
            'OpenDir = "file://10.245.1.18/MIS/DASW/" + DNo.Text

            DAttachfile.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")



            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = dtData.Rows(0).Item("Date")                         'No 
            DDepName.Text = dtData.Rows(0).Item("DepName")                         'No
            DAppName.Text = dtData.Rows(0).Item("AppName")                         'No

            SetFieldData("Job", dtData.Rows(0).Item("Job"))    '處理人員
            SetFieldData("Currency", dtData.Rows(0).Item("Currency"))

            DCustomerCode.Text = dtData.Rows(0).Item("CustomerCode")
            DCustomer.Text = dtData.Rows(0).Item("Customer")
            DVersion.Text = dtData.Rows(0).Item("Version")

            DBuyerCode.Text = dtData.Rows(0).Item("Buyercode")
            DBuyer.Text = dtData.Rows(0).Item("Buyer")
            DAReason.Text = dtData.Rows(0).Item("AReason")

            DASDate.Text = dtData.Rows(0).Item("ASDate")
            DAEDate.Text = dtData.Rows(0).Item("AEDate")

            If dtData.Rows(0).Item("Extend") = "1" Then
                LYearNo.Visible = True
                LYearNo.Text = "延長申請"
                LYearNo.NavigateUrl = "DiscountExtend.aspx?&pNo=" & DNo.Text
            End If

            If dtData.Rows(0).Item("OFormSno") <> "" Then
                ChkYear.Checked = True
                ChkYear.Visible = True
                ChkYear.Enabled = False
                DOFormSno.Text = dtData.Rows(0).Item("OFormSno")
                LNo.Visible = True
                LNo.NavigateUrl = "DiscountSheet_02.aspx?&pFormno=003102&pFormsno=" + Mid(dtData.Rows(0).Item("OFormSno"), 7, 5)

            End If



            '明細資料
            SQL = "Select * From F_DiscountSheetdt "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " order by seqno"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            Dim i As Integer
            Dim ItemCode, ST1, ST2, ST3, ST4, ST5, SIZE, CHAIN, COM, DISCOUNT, wRow As String

            If DBAdapter2.Rows.Count > 0 Then
                For i = 0 To DBAdapter2.Rows.Count - 1

                    ItemCode = "DITEMCODE" + Trim(Str(i))
                    Dim DITEMCODE As TextBox = Me.FindControl(ItemCode)
                    If DBAdapter2.Rows(i).Item("ITEMCODE") <> "" Then
                        DITEMCODE.Text = DBAdapter2.Rows(i).Item("ITEMCODE")

                    End If

                    ST1 = "DST1" + Trim(Str(i))
                    Dim DST1 As DropDownList = Me.FindControl(ST1)

                    If DBAdapter2.Rows(i).Item("ST1") <> "" Then
                        SetFieldData(ST1, DBAdapter2.Rows(i).Item("ST1"))
                    End If

                    ST2 = "DST2" + Trim(Str(i))
                    Dim DST2 As DropDownList = Me.FindControl(ST2)

                    If DBAdapter2.Rows(i).Item("ST2") <> "" Then
                        SetFieldData(ST2, DBAdapter2.Rows(i).Item("ST2"))
                    End If

                    ST3 = "DST3" + Trim(Str(i))
                    Dim DST3 As DropDownList = Me.FindControl(ST3)

                    If DBAdapter2.Rows(i).Item("ST3") <> "" Then
                        SetFieldData(ST3, DBAdapter2.Rows(i).Item("ST3"))
                    End If


                    ST4 = "DST4" + Trim(Str(i))
                    Dim DST4 As DropDownList = Me.FindControl(ST4)

                    If DBAdapter2.Rows(i).Item("ST4") <> "" Then
                        SetFieldData(ST4, DBAdapter2.Rows(i).Item("ST4"))
                    End If



                    ST5 = "DST5" + Trim(Str(i))
                    Dim DST5 As DropDownList = Me.FindControl(ST5)

                    If DBAdapter2.Rows(i).Item("ST5") <> "" Then
                        SetFieldData(ST5, DBAdapter2.Rows(i).Item("ST5"))
                    End If

                    SIZE = "DSIZE" + Trim(Str(i))
                    Dim DSIZE As TextBox = Me.FindControl(SIZE)
                    If DBAdapter2.Rows(i).Item("SIZE") <> "" Then
                        DSIZE.Text = DBAdapter2.Rows(i).Item("SIZE")

                    End If

                    CHAIN = "DCHAIN" + Trim(Str(i))
                    Dim DCHAIN As TextBox = Me.FindControl(CHAIN)
                    If DBAdapter2.Rows(i).Item("CHAIN") <> "" Then
                        DCHAIN.Text = DBAdapter2.Rows(i).Item("CHAIN")

                    End If

                    COM = "DCOM" + Trim(Str(i))
                    Dim DCOM As TextBox = Me.FindControl(COM)
                    If DBAdapter2.Rows(i).Item("COM") <> "" Then
                        DCOM.Text = DBAdapter2.Rows(i).Item("COM")

                    End If


                    DISCOUNT = "DDISCOUNT" + Trim(Str(i))
                    Dim DDISCOUNT As TextBox = Me.FindControl(DISCOUNT)
                    If DBAdapter2.Rows(i).Item("DISCOUNT") <> "" Then
                        DDISCOUNT.Text = DBAdapter2.Rows(i).Item("DISCOUNT")

                    End If



                Next

            End If


        End If
    End Sub
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        Dim idx As Integer = 1
        Dim i As Integer
   
        '幣別
        If pFieldName = "Currency" Then
            DCurrency.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCurrency.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3102'"
                sql = sql & " and dkey = 'Currency'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DCurrency.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCurrency.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        '作業
        If pFieldName = "Job" Then
            DJob.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DJob.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3102'"
                sql = sql & " and dkey = 'Job'"
                sql = sql & " order by data "
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DJob.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DJob.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        '統計區分
        If Mid(pFieldName, 1, 2) = "ST" Then
            Dim Rule As String

            For i = 0 To 9
                Rule = "DST" + Mid(pFieldName, 3, 1) + Trim(Str(i))
                Dim DText As DropDownList = Me.FindControl(Rule)
                DText.Items.Clear()
            Next

            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    For i = 1 To 10
                        Rule = "DST1" + Trim(Str(i))
                        Dim DText As DropDownList = Me.FindControl(Rule)
                        DText.Items.Add(ListItem1)
                    Next

                End If
            Else
                sql = ""
                sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
                sql = sql & " where  cat = '3102'"
                sql = sql & " and dkey = 'ST'"
                If Mid(pFieldName, 3, 1) = "1" Or Mid(pFieldName, 3, 1) = "2" Then
                    sql = sql & "   and data between '0' and '9'"
                End If
                sql = sql & " order by data"


                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

                '新增統計區分
                For Each dtr As Data.DataRow In DBAdapter1.Rows

                    For i = 0 To 9
                        Rule = "DST" + Mid(pFieldName, 3, 1) + Trim(Str(i))
                        Dim DText As DropDownList = Me.FindControl(Rule)
                        If dtr("cno") = "1" Then
                            DText.Items.Clear()
                            DText.Items.Add("")
                        End If
                        DText.Items.Add(dtr("Data"))
                    Next

                Next

            End If
        End If




        If Mid(pFieldName, 1, 3) = "DST" Then
            Dim Rule As String
            Dim COUNTS As Integer


            i = COUNTS
            Rule = pFieldName

            Dim DText As DropDownList = Me.FindControl(Rule)
            DText.Items.Clear()

            sql = ""
            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '3102'"
            sql = sql & " and dkey = 'ST'"
            If Mid(pFieldName, 4, 1) = "1" Or Mid(pFieldName, 4, 1) = "2" Then
                sql = sql & "   and data between '0' and '9'"
            End If
            sql = sql & " order by data"

            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            '新增RULE
            For Each dtr As Data.DataRow In DBAdapter1.Rows

                Rule = pFieldName
                DText.Items.Add(dtr("Data"))
                DText.SelectedValue = pName
            Next


        End If


    End Sub


 
End Class
