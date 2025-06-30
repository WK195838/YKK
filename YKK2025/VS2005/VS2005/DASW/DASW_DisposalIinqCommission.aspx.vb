Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DASW_DisposalIinqCommission
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     �ۭq�[���w
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "DASW_DisposalinqCommission.aspx"
        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetFieldData()
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '�{�b���
        pFormNo = Request.QueryString("pFormNo")    '��渹�X
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub



    '*****************************************************************
    '**
    '**     ��Excel�@�ε{��
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        ' pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '�{���O���P
        DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=DTMW_COMBICommission_ist.xls")     '�{���O���P
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '�{���O���P



    End Sub


    Sub DataList()

      

       
        '���A
        Dim wSts As String = ""
        If DSTS.SelectedValue <> "" Then
            wSts = " and b.Sts in (" + DSTS.Text + ")"
        Else
            wSts = ""
        End If

        '�渹
        Dim wNo As String = ""
        If DNo.Text <> "" Then
            wNo = " and  a.No like '%" + DNo.Text + "%'"
        Else
            wNo = ""
        End If

        '�ӽЪ� 
        Dim wappName As String = ""
        If DAppName.Text <> "" Then
            wappName = " and  appName like '%" + DAppName.Text + "%'"
        Else
            wappName = ""
        End If

        '���o��]
        Dim DISPOSALREASON As String = ""
        If DDISPOSALREASON.SelectedValue <> "" Then
            DISPOSALREASON = " and  DISPOSALREASON like '%" + DDISPOSALREASON.SelectedValue + "%'"
        Else
            DISPOSALREASON = ""
        End If



        '�d������
        Dim DUTYDEPO As String = ""
        If DDUTYDEPO.SelectedValue <> "" Then
            DUTYDEPO = " and  DUTYDEPO like '%" + DDUTYDEPO.SelectedValue + "%'"
        Else
            DUTYDEPO = ""
        End If


        '���o��h
        Dim DISPOSALRULE As String = ""
        If DDISPOSALRULE.SelectedValue <> "" Then
            DISPOSALRULE = " and  DISPOSALRULE like '%" + DDISPOSALRULE.SelectedValue + "%'"
        Else
            DISPOSALRULE = ""
        End If


        '���o����
        Dim PLACE As String = ""
        If DPLACE.SelectedValue <> "" Then
            PLACE = " and  PLACE like '%" + DPLACE.SelectedValue + "%'"
        Else
            PLACE = ""
        End If

        '���o��h
        Dim DISPOSALTYPE As String = ""
        If DDISPOSALTYPE.SelectedValue <> "" Then
            DISPOSALTYPE = " and  DISPOSALTYPE like '%" + DDISPOSALTYPE.SelectedValue + "%'"
        Else
            DISPOSALTYPE = ""
        End If

        '�}�l���
        Dim ASdate As String = ""
        If DASDate.Text <> "" Then
            ASdate = " and  Convert(VARCHAR(10), AppDate, 111) >= '" + DASDate.Text + "'"
        End If

        '�������
        Dim AEdate As String = ""
        If DAEDate.Text <> "" Then
            AEdate = " and  Convert(VARCHAR(10), AppDate, 111) <= '" + DAEDate.Text + "'"
        End If



        Dim i As Integer = 0
        Dim SQL As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "�s��"
        DataGrid1.Columns.Item(1).HeaderText = "�̿બ�A"
        DataGrid1.Columns.Item(2).HeaderText = "�̿���"
        DataGrid1.Columns.Item(3).HeaderText = "�̿��"
        DataGrid1.Columns.Item(4).HeaderText = "�̿ೡ��."
        DataGrid1.Columns.Item(5).HeaderText = "���o�z��"
        DataGrid1.Columns.Item(6).HeaderText = "�d������"
        DataGrid1.Columns.Item(7).HeaderText = "�겣�b�w����R"
        DataGrid1.Columns.Item(8).HeaderText = "���o�ǫh"
        DataGrid1.Columns.Item(9).HeaderText = "���o�~����"

        SQL = " select  a.no as Field1,case when b.sts=0 then '�֩w��' when b.sts =1 then '����' else '����' end as Field2,convert(char(10),appdate,111) as Field3,"
        SQL = SQL + " appname as Field4,deponame as Field5,disposalreason as Field6,dutydepo as Field7,place as Field8,"
        SQL = SQL + " disposalrule as Field9,disposaltype as Field10,"
        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + a.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + a.ApplyID "
        SQL = SQL + " As OPURL,  "
        SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate"
        SQL = SQL + " from (select * from   V_WaitHandle_01 "
        SQL = SQL + "where formno = 6001"
        SQL = SQL + "  and step =1) a,f_disposalsheet b"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'  "
        SQL = SQL + wSts + wNo + wappName + DISPOSALREASON + DUTYDEPO + DISPOSALRULE + PLACE + DISPOSALTYPE + ASdate + AEdate
        SQL = SQL + " order by  a.no desc "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DataGrid1.DataSource = DBAdapter1
        DataGrid1.DataBind()

    End Sub

    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub


    Protected Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub

    Protected Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        '  e.Item.Cells(9).Attributes.Add("style", "vnd.ms-excel.numberformat:@")


        ' e.Item.Cells(11).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        ' e.Item.Cells(12).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        ' e.Item.Cells(13).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        ' e.Item.Cells(14).Attributes.Add("style", "vnd.ms-excel.numberformat:@")

    End Sub


    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData()
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx, i As Integer
        idx = 1


        DSts.Items.Add("")

        '�ӽЭ�]

        DDISPOSALREASON.Items.Clear()
       
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'DISPOSALREASON'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DDISPOSALREASON.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")

            DDISPOSALREASON.Items.Add(ListItem1)
        Next
        dtReferp.Clear()



        '�d������

        DDUTYDEPO.Items.Clear()
     
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'DUTYDEPO'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DDUTYDEPO.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")

            DDUTYDEPO.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()



        '�b�w����

        DPLACE.Items.Clear()
  
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'PLACE'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DPLACE.Items.Add("")
        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data")

            DPLACE.Items.Add(ListItem1)
        Next
        dtReferp2.Clear()



        '���o�ǫh
        DDISPOSALRULE.Items.Clear()
      
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'DISPOSALRULE'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp3 As DataTable = uDataBase.GetDataTable(SQL)
        DDISPOSALRULE.Items.Add("")
        For i = 0 To dtReferp3.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp3.Rows(i).Item("Data")
            ListItem1.Value = dtReferp3.Rows(i).Item("Data")

            DDISPOSALRULE.Items.Add(ListItem1)
        Next
        dtReferp3.Clear()



        '���o�~����

        DDISPOSALTYPE.Items.Clear()
     
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'DISPOSALTYPE'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(SQL)
        DDISPOSALTYPE.Items.Add("")
        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data")

            DDISPOSALTYPE.Items.Add(ListItem1)
        Next
        dtReferp4.Clear()
       
        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"


    End Sub
End Class

