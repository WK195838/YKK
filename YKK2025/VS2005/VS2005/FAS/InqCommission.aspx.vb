Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class InqCommission
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

        Response.Cookies("PGM").Value = "InqCommission.aspx"
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=InqCommission_ist.xls")     '�{���O���P
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

        '�ӽвz��
        Dim TYPE As String = ""
        If DTYPE.SelectedValue <> "" Then
            TYPE = " and  Type like '%" + DTYPE.SelectedValue + "%'"
        Else
            TYPE = ""
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
        DataGrid1.Columns.Item(5).HeaderText = "�ӽвz��"
   

        SQL = " select  a.no as Field1,case when b.sts=0 then '�֩w��' when b.sts =1 then '����' else '����' end as Field2,convert(char(10),appdate,111) as Field3,"
        SQL = SQL + " apper as Field4,b.Division as Field5,Type as Field6,"
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
        SQL = SQL + "where formno = 8001"
        SQL = SQL + "  and step =1) a,f_FASSheet b"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'  "
        SQL = SQL + wSts + wNo + wappName + TYPE + ASdate + AEdate
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


        DSTS.Items.Add("")

        '�ӽЭ�]

        DTYPE.Items.Clear()

        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '8001'"
        SQL = SQL & " and dkey = 'TYPE'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DTYPE.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")

            DTYPE.Items.Add(ListItem1)
        Next
        dtReferp.Clear()




        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"


    End Sub
End Class

