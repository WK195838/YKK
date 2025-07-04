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





Partial Class QAModSheet_02
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As Integer              '預設的控制項位置
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
    Dim oWaves As New WAVES.CommonService
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String
    Dim CheckItemStr, CheckTypeStr, RPReportStr, SampleStr, MaterialStr, LocationStr As String
    Dim NO1 As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        SetParameter()          '設定共用參數
        GetData()

        If Not Me.IsPostBack Then   '不是PostBack

            '加入選項
            CheckItem1()
            NewAttachFilePath()

            ShowFormData()      '顯示表單資料

            NewAttachFilePath2()
        End If



    End Sub
    '---------------------------------------------------
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
        Response.Cookies("PGM").Value = "QAModSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼


    End Sub
    
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()


        Dim SQL, sUrl As String


        SQL = "Select * From F_QAModSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'" '
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then

            DNo.Text = dtData.Rows(0).Item("No")
            DDate.Text = dtData.Rows(0).Item("Date")
            DDepName.Text = dtData.Rows(0).Item("DepName")
            DName.Text = dtData.Rows(0).Item("Name")
            SetFieldData("Checkuser", dtData.Rows(0).Item("CheckUser"))
            DCheckUser.Text = dtData.Rows(0).Item("CheckUser")

            DNO1.Text = dtData.Rows(0).Item("ModNo")
            ModReason.Text = dtData.Rows(0).Item("ModReason")
            ModContent.Text = dtData.Rows(0).Item("ModContent")
            IContent.Text = dtData.Rows(0).Item("IContent")

            ModDate.Text = dtData.Rows(0).Item("ModDate")
            ModNo.Text = dtData.Rows(0).Item("ModNo")
            ModName.Text = dtData.Rows(0).Item("ModName")
            ModDepname.Text = dtData.Rows(0).Item("ModDepName")
            DADate.Text = dtData.Rows(0).Item("ADate")
            DSubject.Text = dtData.Rows(0).Item("Subject")
            DItemName.Text = dtData.Rows(0).Item("ItemName")
            DBackGround.Text = dtData.Rows(0).Item("BackGround")


            DBuyerCode.Text = dtData.Rows(0).Item("BuyerCode")
            DBuyer.Text = dtData.Rows(0).Item("Buyer")


            DCustomerCode.Text = dtData.Rows(0).Item("CustomerCode")
            DCustomer.Text = dtData.Rows(0).Item("Customer")

            If dtData.Rows(0).Item("CheckItem") <> "" Then
                GetCheckItem(dtData.Rows(0).Item("CheckItem"))
            End If


            If dtData.Rows(0).Item("CheckType") <> "" Then
                GetCheckType(dtData.Rows(0).Item("CheckType"))
            End If


            If dtData.Rows(0).Item("RPReport") <> "" Then
                GetRPReport(dtData.Rows(0).Item("RPReport"))
            End If

            If dtData.Rows(0).Item("Sample") <> "" Then
                GetSample(dtData.Rows(0).Item("Sample"))
            End If

            If dtData.Rows(0).Item("Material") <> "" Then
                GetMaterial(dtData.Rows(0).Item("Material"))
            End If


            If dtData.Rows(0).Item("Location") <> "" Then
                GetLocation(dtData.Rows(0).Item("Location"))
            End If


        End If

        DFinishDate.Text = dtData.Rows(0).Item("FinishDate")

        DQCNo.Text = dtData.Rows(0).Item("QCNO")

        If dtData.Rows(0).Item("QCDate") = "1900-01-01 00:00:00.000" Then
            DQCDate.Text = ""
        Else
            DQCDate.Text = dtData.Rows(0).Item("QCDate")
        End If

        SQL = "Select * From F_QASheet "
        SQL &= " Where No =  '" & DNO1.Text & "'" '
        Dim dtData1 As DataTable = uDataBase.GetDataTable(SQL)
        If dtData1.Rows.Count > 0 Then
            sUrl = "QASheet_02.aspx?pFormNo=" & dtData1.Rows(0).Item("FormNo") & "&pFormSno=" & dtData1.Rows(0).Item("FormSno")
            LSheet1.NavigateUrl = sUrl
        End If


        SQL = " select Unique_id,no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2,QCRemark"
        SQL = SQL + " from  F_QAModSheetDT "
        SQL = SQL + "  where No='" + DNo.Text + "'  "
        SQL = SQL + " order by SeqNo "


        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        DCount.Text = "共" + CStr(GridView1.Rows.Count) + "筆"
        If GridView1.Rows.Count > 0 Then

            DHaveData.Text = "1"
        Else
            DHaveData.Text = "0"
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
        SQL = SQL + "Order by  CreateTime Desc, Step Desc, SeqNo Desc "
        GridView2.DataSource = uDataBase.GetDataTable(SQL)
        GridView2.DataBind()

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


        'CheckUser


        sql = "  Select   substring(data,1,CHARINDEX('-', data)-1)appuser,substring(data,CHARINDEX('-', data)+1,len(data)-1)data  from M_referp"
        sql = sql & " where  cat = '8002'"
        sql = sql & " and dkey = 'CheckUser'"
        sql = sql & " and substring(data,1,CHARINDEX('-', data)-1)='" & DName.Text & "'"
        sql = sql & " order by unique_id"

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(sql)
        DCheckUser.Items.Clear()
        DCheckUser.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DCheckUser.Items.Add(ListItem1)
        Next

        dtReferp1.Clear()


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**CheckItem1()
    '**      '加入選項 checkboxlist radioboxlist 
    '**
    '*****************************************************************


    Sub CheckItem1()
        '展開工程選項 
        'Engine-A 單獨選取
        Dim i As Integer

        DCheckItem.RepeatColumns = 1
        DCheckItem.RepeatDirection = RepeatDirection.Horizontal
        DCheckItem.Visible = True
        DCheckItem.Items.Clear()

        Dim SQL As String

        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='CheckItem'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
        CheckItemCount.Text = dt.Rows.Count
        For Each dtr As Data.DataRow In dt.Rows

            Dim CheckItem As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DCheckItem.Items.Add(CheckItem)
        Next

        'Dim SQL As String

        'SQL = "select * from M_referp "
        'SQL = SQL + "where cat = '4001'"
        'SQL = SQL + "and dkey ='EngineSelect-單獨選取'"
        'SQL = SQL + " Order by Unique_ID"

        'Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

        'For Each dtr As Data.DataRow In dt.Rows
        '    Dim EngineA As New ListItem(dtr("Data"), dtr("Data"))
        '    'New ListItem(dtr("RUserName"), dtr("RUserID"))
        '    CheckBoxList1.Items.Add(EngineA)
        'Next


        '營業依賴 2 CheckType


        DCheckType.RepeatColumns = 2
        DCheckType.RepeatDirection = RepeatDirection.Horizontal
        DCheckType.Visible = True
        DCheckType.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='CheckType'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)

        CheckTypeCount.Text = dt1.Rows.Count

        For Each dtr As Data.DataRow In dt1.Rows
            Dim CheckType As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DCheckType.Items.Add(CheckType)
        Next



        '回覆報告書 RPReport


        DRPReport.RepeatColumns = 3
        DRPReport.RepeatDirection = RepeatDirection.Horizontal
        DRPReport.Visible = True
        DRPReport.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='RPReport'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt2 As Data.DataTable = uDataBase.GetDataTable(SQL)
        RPReportCount.Text = dt2.Rows.Count
        For Each dtr As Data.DataRow In dt2.Rows
            Dim RPReport As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DRPReport.Items.Add(RPReport)
        Next




        '樣品 Sample


        DSample.RepeatColumns = 3
        DSample.RepeatDirection = RepeatDirection.Horizontal
        DSample.Visible = True
        DSample.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='Sample'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt3 As Data.DataTable = uDataBase.GetDataTable(SQL)
        SampleCount.Text = dt3.Rows.Count
        For Each dtr As Data.DataRow In dt3.Rows
            Dim Sample As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DSample.Items.Add(Sample)
        Next


        '殘試料 Material


        DMaterial.RepeatColumns = 3
        DMaterial.RepeatDirection = RepeatDirection.Horizontal
        DMaterial.Visible = True
        DMaterial.Items.Clear()





        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='Material'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt4 As Data.DataTable = uDataBase.GetDataTable(SQL)
        MaterialCount.Text = dt4.Rows.Count
        For Each dtr As Data.DataRow In dt4.Rows
            Dim Material As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DMaterial.Items.Add(Material)
        Next


        '配布先 Location


        DLocation.RepeatColumns = 3
        DLocation.RepeatDirection = RepeatDirection.Horizontal
        DLocation.Visible = True
        DLocation.Items.Clear()




        SQL = "select * from M_referp "
        SQL = SQL + "where cat = '8002'"
        SQL = SQL + "and dkey ='Location'"
        SQL = SQL + " Order by Unique_ID"

        Dim dt5 As Data.DataTable = uDataBase.GetDataTable(SQL)
        LocationCount.Text = dt5.Rows.Count
        For Each dtr As Data.DataRow In dt5.Rows
            Dim Location As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DLocation.Items.Add(Location)
        Next


    End Sub


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
    '**CheckResult()
    '**      '組合選項字串 checkboxlist radioboxlist 
    '**
    '*****************************************************************


    Sub CheckResult()
        Dim i, j As Integer
        Dim CheckStr As String
        CheckStr = ""





        j = 1
        For i = 0 To (DCheckItem.Items.Count - 1)
            If (DCheckItem.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DCheckItem.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DCheckItem.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            CheckItemStr = CheckStr

        Next

        CheckStr = ""

        j = 1
        For i = 0 To (DCheckType.Items.Count - 1)
            If (DCheckType.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DCheckType.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DCheckType.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            CheckTypeStr = CheckStr

        Next

        CheckStr = ""
        j = 1
        For i = 0 To (DRPReport.Items.Count - 1)
            If (DRPReport.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DRPReport.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DRPReport.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            RPReportStr = CheckStr

        Next

        CheckStr = ""

        j = 1
        For i = 0 To (DSample.Items.Count - 1)
            If (DSample.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DSample.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DSample.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            SampleStr = CheckStr

        Next


        CheckStr = ""

        j = 1
        For i = 0 To (DMaterial.Items.Count - 1)
            If (DMaterial.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DMaterial.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DMaterial.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            MaterialStr = CheckStr

        Next




        CheckStr = ""
        j = 1
        For i = 0 To (DLocation.Items.Count - 1)
            If (DLocation.Items(i).Selected) Then
                If CheckStr = "" Then
                    CheckStr = Mid(Trim(DLocation.Items(i).Text), 1, 2)
                Else
                    CheckStr = CheckStr + "," + Mid(Trim(DLocation.Items(i).Text), 1, 2)
                End If



                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)

                j = j + 1
            End If
            LocationStr = CheckStr

        Next




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
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If


        If D3.Text = "" Then
            D3.Text = Now.ToString("yyyyMMddHHmmss")
        End If



        OpenDir1 = OpenDir1 + D3.Text + "/核可卡"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + D3.Text + "\核可卡"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If

        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)


        '開啟附檔資料夾路徑
        'DAttachFile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

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
        SQL = SQL + " where cat = '8002'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNo.Text + "/檢測"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text + "\檢測"
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

        '開啟附檔資料夾路徑
        'DAttachFile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 CheckType 字串資料，加入Checked 
    '**
    '*****************************************************************


    Function GetCheckType(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"，"c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(CheckTypeCount.Text) - 1
                If Mid(DCheckType.Items(j).Text, 1, 2) = Result Then
                    DCheckType.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DCheckType.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 CheckItem 字串資料，加入Checked 
    '**
    '*****************************************************************

    Function GetCheckItem(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"，"c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(CheckItemCount.Text) - 1

                If Mid(DCheckItem.Items(j).Text, 1, 2) = Result Then
                    DCheckItem.Items(j).Selected = True

                End If
                If wStep <> 1 And wStep <> 500 Then
                    DCheckItem.Items(j).Enabled = False
                End If


            Next

        Next
        Return Result
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 RPReport 字串資料，加入Checked 
    '**
    '*****************************************************************


    Function GetRPReport(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"，"c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(RPReportCount.Text) - 1

                If Mid(DRPReport.Items(j).Text, 1, 2) = Result Then
                    DRPReport.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DRPReport.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 Sample 字串資料，加入Checked 
    '**
    '*****************************************************************


    Function GetSample(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"，"c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(SampleCount.Text) - 1

                If Mid(DSample.Items(j).Text, 1, 2) = Result Then
                    DSample.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DSample.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**    showform 時將 Material 字串資料，加入Checked 
    '**
    '*****************************************************************

    Function GetMaterial(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"，"c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(MaterialCount.Text) - 1

                If Mid(DMaterial.Items(j).Text, 1, 2) = Result Then
                    DMaterial.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DMaterial.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function

    '*****************************************************************
    '**
    '**    showform 時將Location 字串資料，加入Checked 
    '**
    '*****************************************************************
    Function GetLocation(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {"，"c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(LocationCount.Text) - 1

                If Mid(DLocation.Items(j).Text, 1, 2) = Result Then
                    DLocation.Items(j).Selected = True
                End If
                If wStep <> 1 And wStep <> 500 Then
                    DLocation.Items(j).Enabled = False
                End If

            Next


        Next
        Return Result
    End Function


    Sub GetData()
        Dim SQL As String


        SQL = " select Unique_id,no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Element,Result1,Result2"
        SQL = SQL + " from  F_QAModSheetDT "
        SQL = SQL + "  where  No='" + DTNo.Text + "' "
        SQL = SQL + " order by SeqNo "
        uDataBase.ExecuteNonQuery(SQL)
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            DCount.Text = "共" + CStr(dtData.Rows.Count) + "筆"
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()

        End If


    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub
End Class
