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


Partial Class FinalcheckOriSheet_02
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
    Dim wNo As Integer         '表單流水號
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"
        ' BQCIDATE1.Attributes("onclick") = "calendarPicker('Form1.DAppDate');"
        SetParameter()          '設定共用參數

        If Not Me.IsPostBack Then   '不是PostBack
            NewAttachFilePath()

            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查



            ShowFormData()      '顯示表單資料
            NewAttachFilePath2()

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
        wNo = Request.QueryString("pNo")    '表單號碼
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



        Dim SQL As String
        SQL = "Select * From F_FinalcheckOriSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"

        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DONO.Text = dtData.Rows(0).Item("No")                         'No
            DCUSTOMER.Text = dtData.Rows(0).Item("CUSTOMER")
            SetFieldData("ACCDEPNAME", dtData.Rows(0).Item("ACCDEPNAME"))    '顧客
            SetFieldData("SLDDIVISION", dtData.Rows(0).Item("SLDDIVISION"))
            SetFieldData("RDIVISION", dtData.Rows(0).Item("RDIVISION"))
            DCHECKDATE.Text = dtData.Rows(0).Item("CHECKDATE")
            DORNO.Text = dtData.Rows(0).Item("ORNO")
            DCOLORITEM.Text = dtData.Rows(0).Item("COLORITEM")
            DQTY.Text = dtData.Rows(0).Item("QTY")
            SetFieldData("UNIT1", dtData.Rows(0).Item("UNIT1"))
            DCHECKQTY.Text = dtData.Rows(0).Item("CHECKQTY")
            SetFieldData("UNIT2", dtData.Rows(0).Item("UNIT2"))
            DERRORQTY.Text = dtData.Rows(0).Item("ERRORQTY")
            SetFieldData("UNIT3", dtData.Rows(0).Item("UNIT3"))
            DAPPNAME.Text = dtData.Rows(0).Item("APPNAME")
            DDATE.Text = dtData.Rows(0).Item("DATE")

            DERROR1.Text = dtData.Rows(0).Item("ERROR1")
            DERROR2.Text = dtData.Rows(0).Item("ERROR2")
            DERROR3.Text = dtData.Rows(0).Item("ERROR3")
            DERROR4.Text = dtData.Rows(0).Item("ERROR4")
            DERROR5.Text = dtData.Rows(0).Item("ERROR5")
            DERRORSTS.Text = dtData.Rows(0).Item("ERRORSTS")


            SetFieldData("ECONTENT", dtData.Rows(0).Item("ECONTENT"))
            DECONTENT1.Text = dtData.Rows(0).Item("ECONTENT1")
            SetFieldData("EREASON", dtData.Rows(0).Item("EREASON"))
            DEREASON1.Text = dtData.Rows(0).Item("EREASON1")
            SetFieldData("SITUATION", dtData.Rows(0).Item("SITUATION"))
            DSITUATION1.Text = dtData.Rows(0).Item("SITUATION1")
            SetFieldData("QCANSWER", dtData.Rows(0).Item("QCANSWER"))
            DQCANSWER1.Text = dtData.Rows(0).Item("QCANSWER1")

            DQCI1.Text = dtData.Rows(0).Item("QCI1")
            DQCI2.Text = dtData.Rows(0).Item("QCI2")
            DQCIDATE1.Text = dtData.Rows(0).Item("QCIDATE1")
            DQCIDATE2.Text = dtData.Rows(0).Item("QCIDATE2")
            DQCI3.Text = dtData.Rows(0).Item("QCI3")
            DQCIDATE3.Text = dtData.Rows(0).Item("QCIDATE3")
            DQCIDATE4.Text = dtData.Rows(0).Item("QCIDATE4")

            DQCO1.Text = dtData.Rows(0).Item("QCO1")
            DQCO2.Text = dtData.Rows(0).Item("QCO2")
            DQCODATE1.Text = dtData.Rows(0).Item("QCODATE1")
            DQCODATE2.Text = dtData.Rows(0).Item("QCODATE2")
            DQCO3.Text = dtData.Rows(0).Item("QCO3")
            DQCODATE3.Text = dtData.Rows(0).Item("QCODATE3")
            DQCODATE4.Text = dtData.Rows(0).Item("QCODATE4")


            DFINISHDATE.Text = dtData.Rows(0).Item("FINISHDATE")
            DACCEMPNAME.Text = dtData.Rows(0).Item("ACCEMPNAME")



            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3106'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DONO.Text + "/QC"
            End If

            'Dim OpenDir As String
            'OpenDir = "file://10.245.1.18/MIS/DASW/" + DNo.Text

            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")



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

        '被訴部門
        If pFieldName = "ACCDEPNAME" Then
            DACCDEPNAME.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DACCDEPNAME.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select substring(data,1,CHARINDEX('-', data)-1)data  from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'RDIVISION'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DACCDEPNAME.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DACCDEPNAME.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'SLD工程別        
        If pFieldName = "SLDDIVISION" Then
            DSLDDIVISION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSLDDIVISION.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  substring(data,1,CHARINDEX('-', data)-1)data from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'SLDDIVISION'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSLDDIVISION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSLDDIVISION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '相關部門
        If pFieldName = "RDIVISION" Then
            DRDIVISION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DRDIVISION.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  substring(data,1,CHARINDEX('-', data)-1)data from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'RDIVISION'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DRDIVISION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DRDIVISION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'unit1
        If pFieldName = "UNIT1" Then
            DUNIT1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUNIT1.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'CHECKQTY'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DUNIT1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUNIT1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'UNIT2
        If pFieldName = "UNIT2" Then
            DUNIT2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUNIT2.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'CHECKQTY'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DUNIT2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUNIT2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        'UNIT3
        If pFieldName = "UNIT3" Then
            DUNIT3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DUNIT3.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'CHECKQTY'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DUNIT3.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DUNIT3.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'ECONTENT
        If pFieldName = "ECONTENT" Then
            DECONTENT.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DECONTENT.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'ECONTENT'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DECONTENT.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DECONTENT.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'EREASON
        If pFieldName = "EREASON" Then
            DEREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEREASON.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'EREASON'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DEREASON.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEREASON.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'SITUATION
        If pFieldName = "SITUATION" Then
            DSITUATION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSITUATION.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'SITUATION'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSITUATION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSITUATION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'QCANSWER
        If pFieldName = "QCANSWER" Then
            DQCANSWER.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCANSWER.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3106'"
                sql = sql & " and dkey = 'QCANSWER'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DQCANSWER.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCANSWER.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If






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
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3106'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/QC"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\QC"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If



        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath2()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3106'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + DONO.Text + "/Other"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DONO.Text + "\Other"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If



        '開啟附檔資料夾路徑
        DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub




End Class
