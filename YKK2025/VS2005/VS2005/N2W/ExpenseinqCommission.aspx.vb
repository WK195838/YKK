Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class ExpenseinqCommission
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

        Response.Cookies("PGM").Value = "ExpenseCommission.aspx"
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
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

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

        Response.AppendHeader("Content-Disposition", "attachment;filename=ExpenseCommission_ist.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
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


        Dim DepName As String = ""
        If DDepName.Text <> "" Then
            DDepName.Text = DDepName.Text.Trim(" ")
            DepName = " and DepName = '" + DDepName.Text + "'"
        End If

        Dim AppName As String = ""
        If DAppName.Text <> "" Then
            AppName = " and AppName like '%" + DAppName.Text + "%'"

        End If


        Dim PayMan As String = ""
        If DPayman.SelectedValue <> "" Then
            PayMan = " and PayMan Like'%" + DPayman.SelectedValue + "%'"
        End If



        '進度
        Dim Process As String = ""
        If DProcess.SelectedValue <> "" Then
            If DProcess.SelectedValue = "999" Then
                Process = " and b.Sts in (1,2,3)"
            End If
            Process = " and active =1 and flowtype =1  and step like '%" + DProcess.SelectedValue + "%'"
        Else
            Process = " and (step =1 or step =2) "
        End If

        '單號
        Dim FormNo As String = ""
        Dim No As String = ""
        If DFormNo.Text <> "" Then
            FormNo = " and b.No like '%" + DFormNo.Text + "%'"
        End If

        '開始日期
        Dim AppDateS As Date
        Dim AppDateS1 As String = ""
        If DASDateS.Text <> "" Then
            AppDateS = DASDateS.Text
            AppDateS1 = Format(AppDateS, "yyyy/MM/dd")
            AppDateS1 = " and  Convert(VARCHAR(10), appdate, 111) >= '" + AppDateS1 + "'"
        End If

        '結束日期
        Dim AppDateE As Date
        Dim AppDateE1 As String = ""
        If DAEDateE.Text <> "" Then
            AppDateE = DAEDateE.Text
            AppDateE1 = Format(AppDateE, "yyyy/MM/dd")
            AppDateE1 = " and  Convert(VARCHAR(10), appdate, 111) <= '" + AppDateE1 + "'"
        End If

        '開始日期
        Dim ASdate As Date
        Dim ASdate1 As String = ""
        If DASDate.Text <> "" Then
            ASdate = DASDate.Text
            ASdate1 = Format(ASdate, "yyyy/MM/dd")
            ASdate1 = " and  Convert(VARCHAR(10), date, 111) >= '" + ASdate1 + "'"
        End If

        '結束日期
        Dim AEdate As Date
        Dim AEdate1 As String = ""

        If DAEDate.Text <> "" Then
            AEdate = DAEDate.Text
            AEdate1 = Format(AEdate, "yyyy/MM/dd")
            AEdate1 = " and  Convert(VARCHAR(10), date, 111) <= '" + AEdate1 + "'"
        End If


        Dim i As Integer = 0
        Dim SQL As String = ""
        Dim SQL1 As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "申請日期"
        DataGrid1.Columns.Item(3).HeaderText = "申請者"
        DataGrid1.Columns.Item(4).HeaderText = "部門"
        DataGrid1.Columns.Item(5).HeaderText = "費用負擔"


        SQL = " select  a.no as Field1,case when b.sts=0 then '核定中' when b.sts =1 then '完成' else '取消' end as Field2,convert(char(10),date,111) as Field3,"
        SQL = SQL + " appname as Field4,depname as Field5, payman as Field6, "
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
        SQL = SQL + "where formno = '003105') a,f_ExpenseSheet b "

        'SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and decidehistory like '%" + Request.QueryString("pUserID") + "%'"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno "



        '判斷USERID 是否允許調閱全部人員
        SQL1 = " select data from M_referp"
        SQL1 = SQL1 + "  where cat = 3105 "
        SQL1 = SQL1 + " and dkey = 'reviewid'"
        SQL1 = SQL1 + " and CHARINDEX('" + Request.QueryString("pUserID") + "', data)>0"
        Dim DBReviewid As DataTable = uDataBase.GetDataTable(SQL1)
        If DBReviewid.Rows.Count = 0 Then
            SQL = SQL + " and decidehistory like '%" + Request.QueryString("pUserID") + "%'"
        End If


        'If UCase(Request.QueryString("pUserID")) <> "IT003" And UCase(Request.QueryString("pUserID")) <> "IT013" And UCase(Request.QueryString("pUserID")) <> "IT004" And UCase(Request.QueryString("pUserID")) <> "GAS006" And UCase(Request.QueryString("pUserID")) <> "GAS013" And UCase(Request.QueryString("pUserID")) <> "FAS013" And UCase(Request.QueryString("pUserID")) <> "FAS002" And UCase(Request.QueryString("pUserID")) <> "FAS006" And UCase(Request.QueryString("pUserID")) <> "FAS011" Then
        '    SQL = SQL + " and decidehistory like '%" + Request.QueryString("pUserID") + "%'"
        'End If

        SQL = SQL + wSts + AppDateS1 + AppDateE1 + ASdate1 + AEdate1 + Process + DepName + AppName + PayMan + FormNo
        SQL = SQL + " order by  a.no desc "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DataGrid1.DataSource = DBAdapter1
        DataGrid1.DataBind()
        DCount.Text = "總共" + CStr(DBAdapter1.Rows.Count) + "件"

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


        'Dep

        SQL = "  Select  Unique_ID, Cat, DKey, SUBSTRING(Data, 9,20) Data, CreateUser, CreateTime, ModifyUser, ModifyTime  from M_referp"
        SQL = SQL & " where  cat = '3105'"
        SQL = SQL & " and dkey = 'Dep'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)

        DDepName.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DDepName.Items.Add(ListItem1)
        Next
        dtReferp.Clear()



        'Payman

        SQL = "  Select  *  from M_referp"
        SQL = SQL & " where  cat = '3105'"
        SQL = SQL & " and dkey = 'Payman'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")

            DPayman.Items.Add(ListItem1)

        Next
        dtReferp1.Clear()



        '進度
        DProcess.Items.Clear()

        SQL = " select distinct step,stepname data from  M_Flow "
        SQL = SQL & " where formno = 3105 "
        SQL = SQL & " and step not in (1,999)"
        SQL = SQL & " and flowtype =1 order by step"

        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DProcess.Items.Add("")
        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data")
            ListItem1.Value = dtReferp2.Rows(i).Item("step")

            DProcess.Items.Add(ListItem1)
        Next
        dtReferp2.Clear()


        BASDate1.Attributes("onclick") = "calendarPicker('Form1.DASDateS');"
        BAEDate1.Attributes("onclick") = "calendarPicker('Form1.DAEDateE');"
        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"

    End Sub
End Class

