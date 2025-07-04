Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class NoComminqCommission
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

        Response.Cookies("PGM").Value = "NoComminqCommission.aspx"
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=NoComm_ist.xls")     '程式別不同
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
        If DDivision.Text <> "" Then
            DDivision.Text = DDivision.Text.Trim(" ")
        End If

        Dim EmpName As String = ""
        If DEmpName.Text <> "" Then
            EmpName = " and EmpName like '%" + DEmpName.Text + "%'"

        End If

        Dim Cust As String = ""
        If DCust.Text <> "" Then
            Cust = " and ( Cust like '%" + DEmpName.Text + "%' or CustName like '%" + DEmpName.Text + "%')"

        End If

        Dim SReason As String = ""
        If DSReason.Text <> "" Then
            SReason = " and  SReason = '" + DSReason.SelectedValue + "'"

        End If

        Dim QCNO As String = ""
        If DQCNO.Text <> "" Then
            QCNO = " and  ComplainNo = '" + DQCNO.Text + "'"

        End If

        Dim ORNO As String = ""
        If DORNO.Text <> "" Then
            ORNO = " and  orderno = '" + DORNO.Text + "'"

        End If






        '進度
        Dim Process As String = ""
        If DProcess.SelectedValue <> "" Then
            If DProcess.SelectedValue = "999" Then
                Process = " and b.Sts in (1,2,3)"
            End If
            Process = " and active =1 and flowtype =1  and step like '%" + DProcess.SelectedValue + "%'"
        Else
            Process = " and step =1"
        End If

        '單號
        Dim FormNo As String = ""
        Dim No As String = ""
        If DFormNo.Text <> "" Then
            FormNo = " and b.No like '%" + DFormNo.Text + "%'"
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

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "申請日期"
        DataGrid1.Columns.Item(3).HeaderText = "申請者"
        DataGrid1.Columns.Item(4).HeaderText = "部門"
        DataGrid1.Columns.Item(5).HeaderText = "客戶"
        DataGrid1.Columns.Item(6).HeaderText = "客訴單號"
        DataGrid1.Columns.Item(7).HeaderText = "無償理由"
        DataGrid1.Columns.Item(8).HeaderText = "無償ORNO"


        SQL = " select  a.no as Field1,case when b.sts=0 then '核定中' when b.sts =1 then '完成' else '取消' end as Field2,convert(char(10),date,111) as Field3,"
        SQL = SQL + " empname as Field4,b.division as Field5, Cust+'-'+ CustName as Field6, ComplainNo as Field7,SReason as Field8,OrderNo as Field9, "
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
        SQL = SQL + "where formno = '001172') a"
        ',f_NoCommSheetdt b "
        SQL = SQL + " ,(select distinct a.*,orderno from f_nocommsheet a,f_nocommsheetdt b"
        SQL = SQL + " where a.formsno = b.formsno) b"


        'SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and decidehistory like '%" + Request.QueryString("pUserID") + "%'"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno "
        'If UCase(Request.QueryString("pUserID")) <> "IT003" And UCase(Request.QueryString("pUserID")) <> "IT013" And UCase(Request.QueryString("pUserID")) <> "IT004" And UCase(Request.QueryString("pUserID")) <> "QC007" Then
        '    SQL = SQL + " and decidehistory like '%" + Request.QueryString("pUserID") + "%'"
        'End If

        SQL = SQL + wSts + ASdate1 + AEdate1 + Process + DepName + EmpName + FormNo + ORNO + Cust + SReason + QCNO
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


        'Dep

        SQL = "  Select  Unique_ID, Cat, DKey, SUBSTRING(Data, 9,20) Data, CreateUser, CreateTime, ModifyUser, ModifyTime  from M_referp"
        SQL = SQL & " where  cat = '3105'"
        SQL = SQL & " and dkey = 'Dep'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)

        DDivision.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DDivision.Items.Add(ListItem1)
        Next
        dtReferp.Clear()



        '

        DSReason.Items.Clear()
        
        SQL = "Select  Data  From M_Referp Where Cat='1172' and dkey ='SReason'  Order by DKey, Data "
        Dim dtReasonCode As DataTable = uDataBase.GetDataTable(SQL)

        DSReason.Items.Add("")
        For i = 0 To dtReasonCode.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = Trim(dtReasonCode.Rows(i)("Data"))
            ListItem1.Value = Trim(dtReasonCode.Rows(i)("Data"))
            DSReason.Items.Add(ListItem1)
        Next






        '進度
        DProcess.Items.Clear()

        SQL = " select distinct step,stepname data from  M_Flow "
        SQL = SQL & " where formno = 1172 "
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

        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"

    End Sub
End Class

