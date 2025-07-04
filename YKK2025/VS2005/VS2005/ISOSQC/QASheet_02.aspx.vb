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
Imports System.ComponentModel
Imports System.Windows.Forms





Partial Class QASheet_02
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
    Dim oWaves As New WAVES.CommonService
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String
    Dim CheckItemStr, CheckTypeStr, RPReportStr, SampleStr, MaterialStr, LocationStr As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        SetParameter()          '設定共用參數


        ' GetData()

        '加入選項
        CheckItem1()
        ShowFormData()      '顯示表單資料

        NewAttachFilePath()

        If DQCNo.Text <> "" Then
            DAttachFile2.Visible = True
            NewAttachFilePath2()
        Else
            DAttachFile2.Visible = False
        End If


        'If Not Me.IsPostBack Then   '不是PostBack


        '    '加入選項
        '    CheckItem1()
        '    ShowFormData()      '顯示表單資料

        '    NewAttachFilePath()

        '    If DQCNo.Text <> "" Then
        '        DAttachFile2.Visible = True
        '        NewAttachFilePath2()
        '    Else
        '        DAttachFile2.Visible = False
        '    End If

        'End If
 
    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()

        Try
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
            Response.Cookies("PGM").Value = "QASheet_02.aspx"
            Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
            Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

        Catch ex As Exception
            Response.Write("confirm(‘2023/10/1 以前的資料，請自行調閱！’);")
        End Try


    End Sub
 
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Try
            Dim SQL As String
            SQL = "Select * From F_QASheet "
            SQL &= " Where FormNo =  '" & wFormNo & "'" '
            SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
            Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
            If dtData.Rows.Count > 0 Then

                DNo.Text = dtData.Rows(0).Item("No")                         'No
                DDate.Text = dtData.Rows(0).Item("Date")
                DDepName.Text = dtData.Rows(0).Item("DepName")
                DName.Text = dtData.Rows(0).Item("Name")
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

            SQL = "select  distinct modno from F_QAModSheet"
            SQL = SQL + " where ModNo='" + DNo.Text + "' "
            Dim dtData1 As DataTable = uDataBase.GetDataTable(SQL)
            If dtData1.Rows.Count > 0 Then
                LModNo.Visible = True
                LModNo.Text = "有修改申請"
                '  LTripNo.NavigateUrl = "http://10.245.1.10/N2W/BusinessTripSheet_02.aspx?pFormNo=003115&pFormSno=" + Trim(Ddata.Rows(0).Item("formsno"))
                LModNo.NavigateUrl = "ModNoList.aspx?&pModNo=" + Trim(dtData1.Rows(0).Item("ModNo"))
                'LModNo.NavigateUrl = "ModNoList.aspx?&pModNo=20230811"
            End If




            SQL = " select Unique_id,no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Result1,Result2,QCRemark,OutTestNo,formula "
            SQL = SQL + " from  F_QASheetDT "
            SQL = SQL + "  where  No='" + DNo.Text + "' "
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

        Catch ex As Exception
       
            'uJavaScript.PopMsg(Me, "2023/10/1 以前的資料，請自行調閱！")
            ' Response.Write("<script>alert('2023/10/1 以前的資料，請自行調閱！')</script>")

            'Dim myPath As String = "\\file_server4\Dept Share\QC\EDX\EDX檢測明細\中壢\以前"
            'Dim prc As System.Diagnostics.Process = New System.Diagnostics.Process()
            'prc.StartInfo.FileName = myPath
            'prc.Start()

            Response.Write("<script languge='javascript'>alert('2023/10/1 以前的資料，請自行調閱！');window.open('\\\\file_server4\\Dept Share\\QC\\EDX\\EDX檢測明細\\中壢')</script>")

            '  Response.Write("<script languge='javascript'>alert('2023/10/1 以前的資料，請自行調閱！');window.open('C:\\Users\\mis22')</script>")

            Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
            'IE11
            Response.Write("<script>window.open('', '_self', ''); window.close();</script>")

       

        End Try

     
    End Sub
 


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**CheckItem1()
    '**      '加入選項 checkboxlist radioboxlist 
    '**
    '*****************************************************************

    Sub CheckItem1()
        
        '營業依賴 1 Checkitem

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

            Dim EngineA As New ListItem(dtr("Data"), dtr("Data"))
            'New ListItem(dtr("RUserName"), dtr("RUserID"))
            DCheckItem.Items.Add(EngineA)
        Next
        DCheckItem.Items(2).Selected = True

 


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

        DCheckType.Items(10).Selected = True


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


        DRPReport.Items(1).Selected = True


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

        DSample.Items(1).Selected = True

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

        DMaterial.Items(1).Selected = True


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
        DLocation.Items(0).Selected = True




    End Sub


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
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If





        OpenDir1 = OpenDir1 + DNo.Text + "/核可卡"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text + "\核可卡"
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
        DAttachFile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

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




        OpenDir1 = OpenDir1 + "EDX/" + DQCNo.Text    '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + "EDX\" + DQCNo.Text
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
        DAttachFile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub

    Function GetCheckType(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
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

    Function GetCheckItem(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
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


    Function GetRPReport(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
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



    Function GetSample(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(SampleCount.Text) - 1

                DSample.Items(j).Selected = False

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

    Function GetMaterial(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(MaterialCount.Text) - 1
                DMaterial.Items(j).Selected = False
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


    Function GetLocation(ByVal CheckSTr As String) As String
        Dim str As String = CheckSTr
        Dim sArray As String() = str.Split(New Char() {","c})
        Dim Result As String = ""
        Dim j As Integer
        For Each i As String In sArray
            Result = i.ToString()

            For j = 0 To CInt(LocationCount.Text) - 1
                DLocation.Items(j).Selected = False
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




        SQL = " select Unique_id,no,SeqNo,Type,Supplier,Code,Size,Family,Body,Puller,COLOR,Finish,ApproveNo,OldToNewNo,OldToNew,Remark,Result1,Result2,QCRemark,OutTestNo,formula "
        SQL = SQL + " from  F_QASheetDT "
        SQL = SQL + "  where  No='" + DNo.Text + "' "
        SQL = SQL + " order by SeqNo "

        'uDataBase.ExecuteNonQuery(SQL)
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

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim h1 As New HyperLink
        Dim h2 As New HyperLink
        Dim h3 As New HyperLink
        Dim h4 As New HyperLink
        Dim spec As String = ""

        '增加外注測試連結
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.Cells(18).Text <> "&nbsp;" Then
                h1.Text = e.Row.Cells(18).Text
                ' 連結到待處理LIST
                ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                h1.NavigateUrl = "file://10.245.1.6/wfs$/ISOSQC/008002/OutTest/" + e.Row.Cells(18).Text + ".pdf"
                h1.Target = "_blank"
                e.Row.Cells(18).Text = h1.Text
                e.Row.Cells(18).Controls.Add(h1)

            End If

            '核可卡連結
            If e.Row.Cells(11).Text <> "&nbsp;" Then
                h2.Text = e.Row.Cells(11).Text
                ' 連結到待處理LIST
                h2.NavigateUrl = "http://10.245.1.6/ISOSQC/QASheet_02.aspx?pFormNo=008002&pFormSno=" + Mid(e.Row.Cells(11).Text, 7, 4)

                h2.Target = "_blank"
                e.Row.Cells(11).Text = h2.Text
                e.Row.Cells(11).Controls.Add(h2)

            End If


            '舊換新
            If e.Row.Cells(12).Text <> "&nbsp;" Then
                h3.Text = e.Row.Cells(12).Text
                ' 連結到待處理LIST
                h3.NavigateUrl = "\ISOSQC\QASheet_02.aspx?pFormNo=008002&pFormSno=" + Mid(e.Row.Cells(12).Text, 7, 4)

                h3.Target = "_blank"
                e.Row.Cells(12).Text = h3.Text
                e.Row.Cells(12).Controls.Add(h3)

            End If
         

            '色粉配方連結
            If e.Row.Cells(19).Text = "OK" Then
                h4.Text = e.Row.Cells(19).Text
                ' 連結到待處理LIST
                ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                spec = Trim(e.Row.Cells(5).Text) + " " + Trim(e.Row.Cells(6).Text) + " " + Trim(e.Row.Cells(7).Text) + Trim(e.Row.Cells(8).Text) + Trim(e.Row.Cells(9).Text) + " " + Trim(e.Row.Cells(10).Text)
                h4.NavigateUrl = "file://10.245.1.6/wfs$/ISOSQC/008002/Formula/" + spec
                h4.Target = "_blank"
                e.Row.Cells(19).Text = h4.Text
                e.Row.Cells(19).Controls.Add(h4)

            End If



        End If
    End Sub
End Class
