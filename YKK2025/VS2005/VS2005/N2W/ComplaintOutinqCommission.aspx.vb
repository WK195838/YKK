Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class ComplaintOutinqCommission
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "ComplainOutCommission.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetFieldData()
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '現在日時
        ' pFormNo = Request.QueryString("pFormNo")    '表單號碼
        '   Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub



    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click


        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=ComplainOutCommission_ist.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentEncoding = System.Text.Encoding.Default
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同





    End Sub


    Sub DataList()




        '狀態
        Dim wSts As String = ""
        If DSTS.SelectedValue <> "" Then
            wSts = " and b.Sts in (" + DSTS.Text + ")"
        Else
            wSts = ""
        End If



        Dim APPNAME As String = ""
        If DAPPNAME.Text <> "" Then
            APPNAME = " and APPNAME LIKE '%" + DAPPNAME.Text + "%'"
        End If


        Dim CUSTOMER As String = ""
        If DCUSTOMER.Text <> "" Then
            CUSTOMER = " and CUSTOMER  LIKE '%" + DCUSTOMER.Text + "%'"
        End If


        Dim wstep As String
        wstep = " step =1 "

        '進度
        Dim Process As String = ""
        If DProcess.SelectedValue <> "" Then
            If DProcess.SelectedValue = "999" Then
                Process = " and b.Sts in (1,2,3)"
            End If
            Process = " and active =1 and flowtype <>0  and step ='" + DProcess.SelectedValue + "'"
            wstep = " step =" + DProcess.SelectedValue
        Else
            Process = ""
        End If







        '開始日期
        Dim ASdate As Date
        Dim ASdate1 As String = ""
        If DASDate.Text <> "" Then
            ASdate = DASDate.Text
            ASdate1 = Format(ASdate, "yyyy/MM/dd")
            ASdate1 = " and  Convert(VARCHAR(10), DATE, 111) >= '" + ASdate1 + "'"
        End If

        '結束日期
        Dim AEdate As Date
        Dim AEdate1 As String = ""

        If DAEDate.Text <> "" Then
            AEdate = DAEDate.Text
            AEdate1 = Format(AEdate, "yyyy/MM/dd")
            AEdate1 = " and  Convert(VARCHAR(10),DATE, 111) <= '" + AEdate1 + "'"
        End If

        Dim qcno As String = ""

        If DQCNO.Text <> "" Then
            qcno = " and qcno like '%" + DQCNO.Text + "%'"
        End If

        Dim Accdep1 As String
        If DACCDEP1.SelectedValue = "全部" Then
            Accdep1 = ""
        Else
            Accdep1 = "and accdep1='" + DACCDEP1.SelectedValue + "'"
        End If




        Dim i As Integer = 0
        Dim SQL As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "品管受理編號"
        DataGrid1.Columns.Item(3).HeaderText = "報告日"
        DataGrid1.Columns.Item(4).HeaderText = "顧客名稱"
        DataGrid1.Columns.Item(5).HeaderText = "營業擔當"
        DataGrid1.Columns.Item(6).HeaderText = "責任部門別"
        DataGrid1.Columns.Item(7).HeaderText = "型別"
        DataGrid1.Columns.Item(8).HeaderText = "客訴內容"
        DataGrid1.Columns.Item(9).HeaderText = "OR數量"
        DataGrid1.Columns.Item(10).HeaderText = "客訴數量"
        DataGrid1.Columns.Item(11).HeaderText = "賠償金額"



        SQL = " select  distinct  a.no as Field1,case when b.sts=0 then '核定中' when b.sts =1 then '完成' else '取消' end as Field2, QCNO  as  Field3,convert(char(10),date,111) as Field4,"
        SQL = SQL + " CUSTOMER  as Field5,APPNAME as Field6,Accdep1 as Field7,SPEC as Field8,CPCONTENT as Field9,ORQTY as Field10,CPQTY as Field11,ACOST as Field12, "
        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + a.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + a.ApplyID "
        SQL = SQL + " As OPURL,  "
        SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate"
        SQL = SQL + " from (select * from   V_WaitHandle_01b "
        SQL = SQL + "where formno = 3109 and seqno =1 "
        SQL = SQL + "  and " + wstep + ") a,f_ComplaintOutSheet b"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno  "
        SQL = SQL + wSts + ASdate1 + AEdate1 + Process + APPNAME + CUSTOMER + qcno + Accdep1
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
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx, i As Integer
        idx = 1





        Sql = "  Select  distinct substring(data,1,CHARINDEX('-',data)-1)Data  from M_referp"
        Sql = Sql & " where  cat = '3109'"
        Sql = Sql & " and dkey = 'RDIVISION'  "

        Dim dtReferp As DataTable = uDataBase.GetDataTable(Sql)
        DACCDEP1.Items.Add("全部")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DACCDEP1.Items.Add(ListItem1)
        Next
        dtReferp.Clear()





        '進度
        DProcess.Items.Clear()

        SQL = " select distinct step,stepname data from  M_Flow "
        SQL = SQL & " where formno = 3109 "
        SQL = SQL & " and step not in (1,999)"
        SQL = SQL & " and stepname not like '%cc%'order by step"

        Dim dtReferp3 As DataTable = uDataBase.GetDataTable(SQL)
        DProcess.Items.Add("")
        For i = 0 To dtReferp3.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp3.Rows(i).Item("Data")
            ListItem1.Value = dtReferp3.Rows(i).Item("step")

            DProcess.Items.Add(ListItem1)
        Next
        dtReferp3.Clear()




        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"

    End Sub


End Class

