Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class MPMProcessesReport_02
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
 
    Dim Top As String               '�w�]�������m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wNo As String           '��渹�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim wAgentID As String          '�Q�N�z�HID
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim wNextGate As String         '�U�@���H
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '�x�_��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End

    Dim wUserName As String = ""    '�m�W�N�z��
    Dim HolidayList As New List(Of Integer) '�ΥH�O�����骺�����ޭ�
    Dim indexList As New List(Of Integer)   '�ΥH�O�����ݩ�������������ޭ�
    Dim DateList As New List(Of String)     '�ΥH�O���ҿ�����@�P���

    ''' <summary>
    ''' �H�U���@�Ψ禡���ŧi
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�


        If Not Me.IsPostBack Then   '���OPostBack


            ShowFormData()      '��ܪ����
            UpdateTranFile()    '��s������



        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        wNo = Request.QueryString("pNo")
        'Add-Start by Joy 2009/11/20(2010��ƾ����)
        '�ӽЪ�-�s�զ�ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.Cookies("UserID").Value)
        'Add-End

        Response.Cookies("PGM").Value = "MPMProcessesReport_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��s������
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.Cookies("UserID").Value)
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.Cookies("UserID").Value)
        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�
        'Modify-End
    End Sub



    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("MPMProcessesFilePath")



        SQL = "Select *, case when sts =0 and  datediff(day,getdate(),finishdate) < 0 then  datediff(day,getdate(),finishdate) else 0 end as delayday From F_MPMProcessesSheet "
        SQL = SQL & " Where No =  '" & CStr(wNo) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate           
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DClinter.Text = DBAdapter1.Rows(0).Item("Clinter")              'Clinter
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DDivisionCode.Text = DBAdapter1.Rows(0).Item("DivisionCode")              'Division
            DType1.Text = DBAdapter1.Rows(0).Item("Type1")
            DType2.Text = DBAdapter1.Rows(0).Item("Type2")
            DMapNo.Text = DBAdapter1.Rows(0).Item("MapNo")              'Code
            DQty.Text = DBAdapter1.Rows(0).Item("Qty")              'Qty
            DAQty.Text = DBAdapter1.Rows(0).Item("AQty")              'Qty
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              'MapNo
            DWeight.Text = DBAdapter1.Rows(0).Item("Weight")              'Weight
            DMaterial.Text = DBAdapter1.Rows(0).Item("Material")              'Weight
            DDevNo.Text = DBAdapter1.Rows(0).Item("DevNo")              'Weight
            DManufout.Text = DBAdapter1.Rows(0).Item("Manufout")              'Weight
            LDelayDay.Text = "��ǤѼơG" + Str(DBAdapter1.Rows(0).Item("delayday")) '��ǤѼ�


            If Mid(DBAdapter1.Rows(0).Item("FinishDate").ToString, 1, 4) = "1900" Then
                DFinishDate.Value = ""
            Else
                DFinishDate.Value = DBAdapter1.Rows(0).Item("FinishDate")               '�U��ɶ�
            End If

            If Mid(DBAdapter1.Rows(0).Item("AFinishDate").ToString, 1, 4) = "1900" Then
                DAFinishDate.Value = ""
            Else
                DAFinishDate.Value = DBAdapter1.Rows(0).Item("AFinishDate")               '�U��ɶ�
            End If

            If Trim(DBAdapter1.Rows(0).Item("MapFile")) <> "" Then
                LMapFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  '������
                LMapFile1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  '�~��̿��
                LMapFile.Visible = True
                LMapFile1.Visible = True
            Else
                LMapFile.Visible = False
                LMapFile1.Visible = False
            End If
        End If

        Dim EngineStr, WorkTime As String


        '�����ȥ�ɶ�
        SQL = " select *,"
        SQL = SQL & " case when  startqty <>'' and (leadtime <> '1900-01-01 00:00:00.000'  and alwaysstop is not  null  ) then '�פ�'"
        SQL = SQL & " when  startqty <>'' and  (leadtime <> '1900-01-01 00:00:00.000'  and stopflag =1  )  then '�Ȱ�' "
        SQL = SQL & " when startqty <>'' and  leadhour <> ''  and worktime -leadhour > 0  then  '����:'+convert(char, worktime -leadhour)  else '' end as lead,"
        SQL = SQL & " case when LEnddate is not null then datediff(minute,LEnddate,Startdate)  else '' end as LastWait "
        SQL = SQL & "  from ("
        SQL = SQL & " Select Engine,RecDate,StartDate,EndDate,case when engine not in ('��q(��)','�u��(��)','CNC(��)','CAM(����)','CNC�q��') and  ( (convert(char(8),startdate,108)<'12:00:00' and  convert(char(8),Enddate,108)>'12:45:00' ) "
        SQL = SQL & " or (day(startdate) <> day(enddate) and convert(char(8),StopDate,108)>'12:00:00' and  convert(char(8),StartDate,108)<'12:00:00' )  "
        SQL = SQL & " or (day(startdate) <> day(enddate) and convert(char(8),Enddate,108)>'12:45:00' ))  "
        SQL = SQL & "  then WorkTime else Worktime end as WorkTime,Starter,remark,SeqNo,StartQty,EndQty,Stoptime,engineno "
        SQL = SQL & " ,case when leadtime='1900-01-01 00:00:00.000'  then '�L�]�w' else   CONVERT(char(19), leadtime, 120)  end as leadtime,leadhour*60 leadhour,alwaysstop,stopflag,"
        SQL = SQL & "  (   select b.enddate  from F_MPMProcessesSheetDT as b where b.seqno = ( select max(c.seqno+c.seqno1) from F_MPMProcessesSheetDT as c where c.seqno < a.seqno  "
        SQL = SQL & "  and  delmark =0 and No =  '" & CStr(wNo) & "'  )  And delmark = 0 And No =  '" & CStr(wNo) & "'  ) as LEnddate  "
        SQL = SQL & "  From F_MPMProcessesSheetDT a"
        SQL = SQL & " Where delmark =0 and No =  '" & CStr(wNo) & "'  )a order by SeqNo"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

        For Each dtr As Data.DataRow In dt.Rows
            EngineStr = "DEngine" + Trim(Str(dtr("SeqNo")))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = dtr("Engine")
        Next
        SQL = "select seqno,sum(isnull(convert(int,worktime-distime),0))worktime   from ("
        SQL = SQL & " select   seqno,case when engine not in ('��q(��)','�u��(��)','CNC(��)','CAM(����)','CNC�q��') and  "
        SQL = SQL & " ( (convert(char(8),startdate,108)<'12:00:00'  and  convert(char(8),Enddate,108)>'12:45:00' ) "
        SQL = SQL & "  or (day(startdate) <> day(enddate) and convert(char(8),StartDate,108)<'12:00:00' ) or "
        SQL = SQL & " (day(startdate) <> day(enddate) and convert(char(8),Enddate,108)>'12:45:00' ))"
        SQL = SQL & " then 0 else 0 end as distime, worktime  from F_MPMProcessesSheetdt "
        SQL = SQL & " Where delmark =0 and No =  '" & CStr(wNo) & "' "
        SQL = SQL & " )a group by  seqno order by SeqNo"
        Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)

        For Each dtr As Data.DataRow In dt1.Rows
            WorkTime = "DWorktime" + Trim(Str(dtr("SeqNo")))
            Dim DText1 As TextBox = Me.FindControl(WorkTime)
            If dtr("Worktime") = 0 Then
                DText1.Text = ""
            Else
                DText1.Text = dtr("Worktime")
            End If

        Next

        GridView1.DataSource = dt
        GridView1.DataBind()


    End Sub



    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim Enginecode As String
        Enginecode = e.Row.Cells(0).Text

        If Enginecode = "�u��(��)" Or Enginecode = "��q(��)" Or Enginecode = "CNC(��)" Or Enginecode = "CAM(����)" Then
            e.Row.Cells(0).ForeColor = Drawing.Color.Red
            e.Row.Cells(1).ForeColor = Drawing.Color.Red
            e.Row.Cells(2).ForeColor = Drawing.Color.Red
            e.Row.Cells(3).ForeColor = Drawing.Color.Red
            e.Row.Cells(5).ForeColor = Drawing.Color.Red
            e.Row.Cells(6).ForeColor = Drawing.Color.Red
        End If

    End Sub

End Class

