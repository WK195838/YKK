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


Partial Class ComplaintOutSheet_02
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
        'ShowFormData()      '顯示表單資料

        'NewAttachFilePath()
        'NewAttachFilePath2()
        'NewAttachFilePath3()
        'NewAttachFilePath4()
        'NewAttachFilePath5()
        'Total()


        If Not Me.IsPostBack Then   '不是PostBack
            ShowFormData()      '顯示表單資料

            Total()
            NewAttachFilePath()
            NewAttachFilePath2()
            NewAttachFilePath3()
            NewAttachFilePath4()
            NewAttachFilePath5()
            NewAttachFilePath6()


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

        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("ComplaintOutPath")


        Dim SQL As String
        SQL = "Select * From F_ComplaintOutSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
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

            DTYPE1.Text = dtData.Rows(0).Item("TYPE1")
            SetFieldData("TYPE2", dtData.Rows(0).Item("TYPE2"))



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

            DMFGDATE.Text = dtData.Rows(0).Item("MFGDATE")


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

            SetFieldData("QCDESCTYPE", dtData.Rows(0).Item("QCDESCTYPE"))
        
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
                DFINALDATE.Text = ""
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

            If Mid(dtData.Rows(0).Item("MFGDATE").ToString, 1, 4) = "1900" Then
                DMFGDATE.Text = ""
            Else
                DMFGDATE.Text = dtData.Rows(0).Item("MFGDATE")
            End If


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
            OpenDir2 = "file://" + DBAdapter3.Rows(0).Item("Data") + DNO.Text + "/營業"

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
                sql = sql & " and data like '%" + Trim(DACCDEP1.SelectedValue) + "%') order by unique_id "


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
                sql = "  select   distinct data  from  M_referp where cat ='3109' and left(Dkey,4) ='What'"
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
        SQL = SQL + " where cat = '3109'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


 



        OpenDir1 = OpenDir1 + DNO.Text + "/營業"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNO.Text + "\營業"
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
            DChkData1.Checked = True

        Else
            DChkData1.Checked = False

        End If


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

        Else
            DChkData2.Checked = False

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

        Else
            DChkData3.Checked = False

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
            '   DChkData2.Text = Str(FileCount) + "件"
        Else
            DChkData4.Checked = False
            ' DChkData2.Text = ""
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
            '   DChkData2.Text = Str(FileCount) + "件"
        Else
            DChkData5.Checked = False
            ' DChkData2.Text = ""
        End If

        '開啟附檔資料夾路徑
        DAttachfile5.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub

    Sub Total()
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim acost, bcost, ccost, dcost, ecost, allcost, HOUR, d1, d2, d3, d4, d5 As Integer
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

            allcost = acost + bcost + ccost + dcost + ecost
            DALLCOST.Text = Str(allcost)

            If CInt(DALLCOST.Text) <= 30000 Then
                DPOINT2.Text = "0"
            ElseIf CInt(DALLCOST.Text) > 30001 And CInt(DALLCOST.Text) <= 300000 Then
                DPOINT2.Text = "1"
            ElseIf CInt(DALLCOST.Text) > 300001 And CInt(DALLCOST.Text) <= 1000000 Then
                DPOINT2.Text = "3"
            ElseIf CInt(DALLCOST.Text) > 1000001 And CInt(DALLCOST.Text) <= 3000000 Then
                DPOINT2.Text = "5"
            ElseIf CInt(DALLCOST.Text) > 3000001 Then
                DPOINT2.Text = "10"
            End If


        Catch ex As Exception
            ' uJavaScript.PopMsg(Me, "請輸入數字")
        End Try

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


 
End Class
