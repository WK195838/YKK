Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorinqPriority
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
    Dim sts As String





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "DTMW_NewColorinqPriority.aspx"
        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetSearchItem()
            SetFieldData()
            DataList()
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


    Sub SetSearchItem()
 
 

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

        Response.AppendHeader("Content-Disposition", "attachment;filename=DTM_Priority.xls")     '�{���O���P
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


        'Dim ctl As System.Web.UI.Control = Me.DataGrid1
        ''DataGrid1�O�A�b���餤��񪺱���
        'HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=Excel.xls")
        '' HttpContext.Current.Response.Charset = "UTF-8"
        'HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.Default
        'HttpContext.Current.Response.ContentType = "application/ms-excel"
        'ctl.Page.EnableViewState = False
        'Dim tw As System.IO.StringWriter = New System.IO.StringWriter()
        'Dim hw As System.Web.UI.HtmlTextWriter = New System.Web.UI.HtmlTextWriter(tw)
        'ctl.RenderControl(hw)
        'HttpContext.Current.Response.Write(tw.ToString())
        'HttpContext.Current.Response.End()



    End Sub


    Sub DataList()
        sts = " and b.sts =0"
        Dim TABLE As String
       

        TABLE = "V_WaitHandle_01"
        Dim i As Integer = 0
        Dim SQL As String

        Dim pForm As String = ""
        If DForm.SelectedValue <> "" Then
            pForm = " and a.formno ='" + DForm.SelectedValue + "'"
        End If

        Dim pName As String = ""
        If DName.Text <> "" Then
            pName = " and a.ApplyName like '%" + DName.Text + "%'"
        End If


        Dim pCustomer As String = ""
        If DCustomer.Text <> "" Then
            pCustomer = " and b.Customer like '%" + DCustomer.Text + "%'"
        End If

        Dim pBuyer As String = ""
        If DBuyer.Text <> "" Then
            pBuyer = " and b.Buyer like '%" + DBuyer.Text + "%'"
        End If

        Dim pPriority As String = ""
        If DPriority.SelectedValue <> "" Then
            pPriority = " where  PriorityNo ='" + DPriority.SelectedValue + "'"
        End If



        DataGrid1.Columns.Item(0).HeaderText = "�s��"
        DataGrid1.Columns.Item(1).HeaderText = "�̿બ�A"
        DataGrid1.Columns.Item(2).HeaderText = "�̿���"
        DataGrid1.Columns.Item(3).HeaderText = "������"

        DataGrid1.Columns.Item(4).HeaderText = "�̿���"
        DataGrid1.Columns.Item(5).HeaderText = "�̿ೡ��"
        DataGrid1.Columns.Item(6).HeaderText = "�̿��"
        DataGrid1.Columns.Item(7).HeaderText = "�Ȥ�"
        DataGrid1.Columns.Item(8).HeaderText = "BUYER"

        DataGrid1.Columns.Item(9).HeaderText = "�Ȥ��W"
        DataGrid1.Columns.Item(10).HeaderText = "�Ȥ�⸹"
        DataGrid1.Columns.Item(11).HeaderText = "���~YKK�⸹"
        DataGrid1.Columns.Item(12).HeaderText = "PANTONE�⸹"

        DataGrid1.Columns.Item(13).HeaderText = "YKK��O"
        DataGrid1.Columns.Item(14).HeaderText = "YKK�⸹"
        DataGrid1.Columns.Item(15).HeaderText = "�ݥΦ���Y"


        DataGrid1.Columns.Item(16).HeaderText = "�ݥΦ�VF�W/�U��"

        DataGrid1.Columns.Item(17).HeaderText = "�ݥΦ�VF����"
        DataGrid1.Columns.Item(18).HeaderText = "�s�¦�"


        DataGrid1.Columns.Item(19).HeaderText = "����/�^�����⸹"



        SQL = " select  *,Priority from ( "
        SQL = SQL + " SELECT a.*,isnull(Priority,'')as Priority,isnull(PriorityNo,9) as PriorityNo FROM ("
        SQL = SQL + "SELECT "
        SQL = SQL + " b.No  As Field1,"
        SQL = SQL + " case b.Sts When '0' Then '�֩w��' When '1' Then '����' When '2' Then '����' Else '����' End As Field2, "
        SQL = SQL + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3, "
        SQL = SQL + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate, "
        SQL = SQL + " formname  as Field4,"
        SQL = SQL + " a.Divname As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"

        SQL = SQL + " YKKColorType As Field9,YKKColorCode as Field10, case  when a.formno ='005011' then  YKKColorCodeSLD else sldcolor end As Field11,"
        SQL = SQL + " case when a.formno ='005011' then  VFCOLOR else VFCOLOR end As Field12,case when a.formno ='005011' then ykkcolorcodevf else vfcolor end as Field13,NewOldColor,"


        SQL = SQL + " '....' as WorkFlow, ViewURL, "


        SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + a.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + a.ApplyID "
        SQL = SQL + " As OPURL,  "


        SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,customerColorCode,overSeaYkkCode,pantonecode,ReColorCode"
        SQL = SQL + " from " + TABLE + " a,V_NewColor b"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'" + sts + pForm + pName + pCustomer + pBuyer
        SQL = SQL + " )A left join F_DTMWPriority b on a.Field1 =b.No   "
        SQL = SQL + " )a " + pPriority
        SQL = SQL + "  order by  PriorityNo, Field3 desc "



        '   SQL1 = " SELECT * FROM (" + SQL + ")A WHERE 1=1 "
        '  SQL1 = SQL1 + AFSdate1 + AFEdate1
        ' SQL1 = SQL1 + " order by Field1 DESC  "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        DataGrid1.Visible = True
        DataGrid1.DataSource = DBAdapter1
        DataGrid1.DataBind()
        If DataGrid1.Items.Count > 0 Then
            BExcel.Visible = True
        Else
            BExcel.Visible = False

        End If






    End Sub

    Protected Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub

    Protected Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound



        e.Item.Cells(10).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(11).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(12).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(13).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(14).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(15).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        'e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        'e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")

    End Sub

 

    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData()
        Dim DBDataSet1 As New DataSet

        Dim SQL As String

        Dim i As Integer


        SQL = "   select formno,formname as Data   from M_form"
        SQL = SQL & " where formno between '005001' and '005099'"
        SQL = SQL & " and formno <> '005007'"
        SQL = SQL & " and active<>0 order by formname "


        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DForm.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("formno")

            DForm.Items.Add(ListItem1)
        Next
        dtReferp.Clear()



    End Sub

    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataList()
    End Sub

 
End Class

