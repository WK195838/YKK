Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DiscountinqCommission
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

        Response.Cookies("PGM").Value = "DiscountinqCommission.aspx"
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=ISMSCommission_ist.xls")     '程式別不同
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


        Dim Appname As String = ""
        If DAppName.Text <> "" Then
            Appname = " and appname  like '%" + DAppName.Text + "%'"
        End If


        Dim Buyer As String = ""
        If DBuyer.Text <> "" Then
            Buyer = " and (B.Buyer  like '%" + DBuyer.Text + "%' or buyercode like '%" + DBuyer.Text + "%')"
        End If


        Dim Customer As String = ""
        If DCustomer.Text <> "" Then
            Customer = " and (Customer like '%" + DCustomer.Text + "%' or customercode like '%" + DCustomer.Text + "%')"
        End If


        Dim Version As String = ""
        If DVersion.Text <> "" Then
            Version = " and Version like '%" + DVersion.Text + "%'"
        End If



        '單號
        Dim NO As String = ""
        If DNO.Text <> "" Then
            NO = " and  a.NO like '%" + DNO.Text + "%'"
        Else
            NO = ""
        End If

      

        '處理人員
        Dim Job As String = ""
        If DJob.SelectedValue <> "" Then
            Job = " and  job like '%" + DJob.SelectedValue + "%'"
        Else
            Job = ""
        End If



        '進度
        Dim Process As String = ""
        If DProcess.SelectedValue <> "" Then
            If DProcess.SelectedValue = "999" Then
                Process = " and b.Sts in (1,2,3)"
            End If
            Process = " and active =1 and flowtype <>0  and step like '%" + DProcess.SelectedValue + "%'"
        Else
            Process = ""
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
            AEdate1 = " and  Convert(VARCHAR(10),date, 111) <= '" + AEdate1 + "'"
        End If

        Dim Itemcode As String = ""

        If DITEMCODE.Text <> "" Then
            Itemcode = "and Itemcode = '" + DITEMCODE.Text + "'"

        End If

        Dim wTable As String = ""
        If DTable.SelectedValue = "未封存" Then
            wTable = " V_WaitHandle_01"
        Else
            wTable = " V_WaitHandle_old_01"
        End If



        Dim i As Integer = 0
        Dim SQL As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "申請日期"
        DataGrid1.Columns.Item(3).HeaderText = "申請者"
        DataGrid1.Columns.Item(4).HeaderText = "客戶"
        DataGrid1.Columns.Item(5).HeaderText = "BUYER"
        DataGrid1.Columns.Item(6).HeaderText = "單價版本"
        DataGrid1.Columns.Item(7).HeaderText = "ITEM CODE"

        SQL = " select* from ("
        SQL = SQL + " select a.*,isnull(b.itemcode,'') as Field8 from ( "
        SQL = SQL + " select  a.no as Field1,case when b.sts=0 then '核定中' when b.sts =1 then '完成' else '取消' end as Field2,convert(char(10),date,111) as Field3,"
        SQL = SQL + " appname as Field4,customer as Field5, b.BUYER as Field6,Version as Field7,"
        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + a.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + a.ApplyID "
        SQL = SQL + " As OPURL,  "
        SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate ,a.formsno,a.createtime "
        SQL = SQL + " from (select * from  " + wTable
        SQL = SQL + " where formno = '003102'"
        SQL = SQL + "  and step in (1,2)"
        SQL = SQL + "  )"
        SQL = SQL + "  a,f_DiscountSheet b   "
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno   "
        SQL = SQL + wSts + Job + ASdate1 + AEdate1 + Process + Appname + Buyer + Customer + Version + NO + Itemcode
        SQL = SQL + " )a left join f_DiscountSheetdt b"
        SQL = SQL + " on a.formsno =b.formsno  "
        SQL = SQL + "  )a order by  createtime desc ,Formsno desc"
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

        '作業

        DJob.Items.Clear()

        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '3102'"
        SQL = SQL & " and dkey = 'Job'"
        SQL = SQL & " order by data "

        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DJob.Items.Add("")
        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data")

            DJob.Items.Add(ListItem1)
        Next
        dtReferp2.Clear()



        '進度
        DProcess.Items.Clear()

        SQL = " select distinct step,stepname data from  M_Flow "
        SQL = SQL & " where formno = 3102 "
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

