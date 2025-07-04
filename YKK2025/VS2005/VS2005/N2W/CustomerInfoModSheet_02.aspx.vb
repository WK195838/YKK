Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic


Partial Class CustomerInfoModSheet_02
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
        SetParameter()
        ShowFormData()
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
        Response.Cookies("PGM").Value = "CustomerInfoModSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼


    End Sub


    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                             System.Configuration.ConfigurationManager.AppSettings("CustomerInfoModiPath")  'WIS-TempPath

        Dim SQL As String
        SQL = "Select * From F_CustomerInfoModSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCust As DataTable = uDataBase.GetDataTable(SQL)
        If dtCust.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtCust.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = dtCust.Rows(0).Item("Date")                         'No 
            DDepName.Text = dtCust.Rows(0).Item("DepName")                         'No
            DAppName.Text = dtCust.Rows(0).Item("AppName")                         'No
            DCustomerCode.Text = dtCust.Rows(0).Item("CustomerCode")                         'No
            DCUSTOMER.Text = dtCust.Rows(0).Item("Customer")
            DRelationCode.Text = dtCust.Rows(0).Item("RelationCode")

            SetFieldData("Location", dtCust.Rows(0).Item("Location"))    ' 

            SetFieldData("Job", dtCust.Rows(0).Item("Job"))    ' 


            DIDNumber.Text = dtCust.Rows(0).Item("IDNumber")                         'No
            DTEL1.Text = dtCust.Rows(0).Item("Tel1")                         'No

            DFAX1.Text = dtCust.Rows(0).Item("Fax1")                         'No


            SetFieldData("Sales", dtCust.Rows(0).Item("Sales"))    ' 
            SetFieldData("Goods", dtCust.Rows(0).Item("Goods"))    ' 
            SetFieldData("Delivery", dtCust.Rows(0).Item("Delivery"))    ' 

            DCustoms.Text = dtCust.Rows(0).Item("Customs")                         'No

            DNameCH.Text = dtCust.Rows(0).Item("NameCH")                         'No
            DNameEN.Text = dtCust.Rows(0).Item("NameEN")
            DInvoiceCH.Text = dtCust.Rows(0).Item("InvoiceCH")
            DInvoiceEN.Text = dtCust.Rows(0).Item("InvoiceEN")
            DPostCode.Text = dtCust.Rows(0).Item("PostCode")
            DTJCode.Text = dtCust.Rows(0).Item("TJCode")

            DAddCH.Text = dtCust.Rows(0).Item("AddCH")                         'No
            DAddEN.Text = dtCust.Rows(0).Item("AddEN")
            DRemark.Text = dtCust.Rows(0).Item("Remark")
            DPaytype.Text = dtCust.Rows(0).Item("Paytype")

            If dtCust.Rows(0).Item("AttachFile") <> "" Then
                LAttachfile.NavigateUrl = Path & dtCust.Rows(0).Item("AttachFile") '折扣
                LAttachfile.Visible = True
            Else
                LAttachfile.Visible = False
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
        '擔當者及部門 


        'Goods
        If pFieldName = "Goods" Then
            DGoods.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DGoods.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Goods'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)

                DGoods.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DGoods.Items.Add(ListItem1)



                Next
                dtReferp.Clear()
            End If
        End If


        'Sales
        If pFieldName = "Sales" Then
            DSales.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSales.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Sales'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSales.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSales.Items.Add(ListItem1)



                Next
                dtReferp.Clear()
            End If
        End If



        'Location
        If pFieldName = "Location" Then
            DLocation.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLocation.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Location'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLocation.Items.Add(ListItem1)



                Next
                dtReferp.Clear()
            End If
        End If

        'Delivery
        If pFieldName = "Delivery" Then
            DDelivery.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDelivery.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Delivery'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDelivery.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'Job
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
                sql = sql & " where  cat = '3103'"
                sql = sql & " and dkey = 'Job'"
                sql = sql & " order by data"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
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



    End Sub


End Class
