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


Partial Class ComplaintInSheet_02
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"
        ' BQCIDATE1.Attributes("onclick") = "calendarPicker('Form1.DAppDate');"
        SetParameter()          '設定共用參數

        If Not Me.IsPostBack Then   '不是PostBack
            NewAttachFilePath()

            ShowFormData()      '顯示表單資料
            NewAttachFilePath2()
            NewAttachFilePath3()



        Else

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
        Response.Cookies("PGM").Value = "ComplaintInSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼


    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                            CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("Http") & _
                    System.Configuration.ConfigurationManager.AppSettings("ComplaintInPath")
        Dim FileName As String = ""

        Dim SQL As String
        SQL = "Select * From F_ComplaintInSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNO.Text = dtData.Rows(0).Item("No")                         'No




            DCOMDATE.Text = dtData.Rows(0).Item("COMDATE")
            SetFieldData("COMDEPNAME", dtData.Rows(0).Item("COMDEPNAME"))
            DCOMNAME.Text = dtData.Rows(0).Item("COMNAME")
            DCOMADMIN.Text = dtData.Rows(0).Item("COMADMIN")

            DCUSTOMERCODE.Text = dtData.Rows(0).Item("CUSTOMERCODE")
            DCUSTOMER.Text = dtData.Rows(0).Item("CUSTOMER")
            DBUYERCODE.Text = dtData.Rows(0).Item("BUYERCODE")
            DBUYER.Text = dtData.Rows(0).Item("BUYER")
            DORNO.Text = dtData.Rows(0).Item("ORNO")
            DColorCode.Text = dtData.Rows(0).Item("ColorCode")

            DREMARK3.Text = dtData.Rows(0).Item("REMARK3")
            DSPEC.Text = dtData.Rows(0).Item("SPEC")
            DSPECNAME.Text = dtData.Rows(0).Item("SPECNAME")
            DORQTY.Text = dtData.Rows(0).Item("ORQTY")
            SetFieldData("UNIT1", dtData.Rows(0).Item("UNIT1"))
            DMATERIALNO.Text = dtData.Rows(0).Item("MATERIALNO")
            DREMARK4.Text = dtData.Rows(0).Item("REMARK4")
            DSTOCKSTS.Text = dtData.Rows(0).Item("STOCKSTS")
            SetFieldData("STOCKCHECK", dtData.Rows(0).Item("STOCKCHECK"))
            SetFieldData("COMINTYPE", dtData.Rows(0).Item("COMINTYPE"))


            DSTATUS.Text = dtData.Rows(0).Item("STATUS")

            If Mid(dtData.Rows(0).Item("DELIVERYDATE").ToString, 1, 4) = "1900" Then
                DDELIVERYDATE.Text = ""
            Else
                DDELIVERYDATE.Text = dtData.Rows(0).Item("DELIVERYDATE")
            End If
            SetFieldData("LOCATION", dtData.Rows(0).Item("LOCATION"))
            SetFieldData("LASTQC", dtData.Rows(0).Item("LASTQC"))
            SetFieldData("StstusList", dtData.Rows(0).Item("StstusList"))

            DQCNO.Text = dtData.Rows(0).Item("QCNO")
            SetFieldData("ACCDEPNAME", dtData.Rows(0).Item("ACCDEPNAME"))
            DACCEMPNAME.Text = dtData.Rows(0).Item("ACCEMPNAME")
            DREMARK1.Text = dtData.Rows(0).Item("REMARK1")

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

            SetFieldData("FCONFIRM", dtData.Rows(0).Item("FCONFIRM"))
            SetFieldData("FCONTENT", dtData.Rows(0).Item("FCONTENT"))
            SetFieldData("FREASON", dtData.Rows(0).Item("FREASON"))
            SetFieldData("SITUATION", dtData.Rows(0).Item("SITUATION"))
            SetFieldData("ANSWER", dtData.Rows(0).Item("ANSWER"))
            DERRORQTY.Text = dtData.Rows(0).Item("ERRORQTY")
            DERRORP.Text = dtData.Rows(0).Item("ERRORP")

            SetFieldData("SHIP", dtData.Rows(0).Item("SHIP"))
            SetFieldData("SLDDIVISION", dtData.Rows(0).Item("SLDDIVISION"))
            SetFieldData("UNIT2", dtData.Rows(0).Item("UNIT2"))
            DREMARK2.Text = dtData.Rows(0).Item("REMARK2")
            DBUQTY.Text = dtData.Rows(0).Item("BUQTY")


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
            DQCDate.Text = dtData.Rows(0).Item("QCDate")

            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3108'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNO.Text + "/投訴部門"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "\\" + DBAdapter3.Rows(0).Item("Data1") + DNO.Text + "\投訴部門"

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
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMDEPNAME'"
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

        '投訴部門
        If pFieldName = "COMDEPNAME" Then
            DCOMDEPNAME.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCOMDEPNAME.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMDEPNAME'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DCOMDEPNAME.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCOMDEPNAME.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        '投訴部門
        If pFieldName = "LASTQC" Then
            DLASTQC.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLASTQC.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'LASTQC'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DLASTQC.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLASTQC.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        '投訴狀況
        If pFieldName = "StstusList" Then
            DStstusList.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DStstusList.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  data  from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'statuslist'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DStstusList.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DStstusList.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If






        'STOCKCHECK
        If pFieldName = "STOCKCHECK" Then
            DSTOCKCHECK.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSTOCKCHECK.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'STOCKCHECK'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSTOCKCHECK.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSTOCKCHECK.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'COMINTYPE
        If pFieldName = "COMINTYPE" Then
            DCOMINTYPE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCOMINTYPE.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'COMINTYPE'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DCOMINTYPE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCOMINTYPE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If





        'LOCATION
        If pFieldName = "LOCATION" Then
            DLOCATION.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLOCATION.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'LOCATION'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DLOCATION.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLOCATION.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'FREASON
        If pFieldName = "FREASON" Then
            DFREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFREASON.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'FREASON'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DFREASON.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFREASON.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'SHIP
        If pFieldName = "SHIP" Then
            DSHIP.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSHIP.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'SHIP'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DSHIP.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSHIP.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If




        'FCONTENT
        If pFieldName = "FCONTENT" Then
            DFCONTENT.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFCONTENT.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'FCONTENT'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DFCONTENT.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFCONTENT.Items.Add(ListItem1)
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
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'SITUATION'"
                sql = sql & " order by unique_id"
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





        'CONFIRM2
        If pFieldName = "FCONFIRM" Then
            DFCONFIRM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFCONFIRM.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'CONFIRM2'"
                sql = sql & " order by DATA"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DFCONFIRM.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFCONFIRM.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        'QCANSWER
        If pFieldName = "ANSWER" Then
            DANSWER.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DANSWER.Items.Add(ListItem1)
                End If
            Else

                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3108'"
                sql = sql & " and dkey = 'QCANSWER'"
                sql = sql & " order by unique_id"
                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DANSWER.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DANSWER.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
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
        SQL = SQL + " where cat = '3108'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If



        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/投訴部門"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\投訴部門"
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
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath2()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3108'"
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
        SQL = SQL + " where cat = '3108'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNO.Text + "/被訴部門"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\被訴部門"
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
            '   DChkData2.Text = Str(FileCount) + "件"
        Else
            DChkData3.Checked = False
            ' DChkData2.Text = ""
        End If
        '開啟附檔資料夾路徑
        DAttachfile3.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub

End Class
