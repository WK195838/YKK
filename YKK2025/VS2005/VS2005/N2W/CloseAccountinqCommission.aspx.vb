Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class CloseAccountinqCommission
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

        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"


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

        Response.AppendHeader("Content-Disposition", "attachment;filename=CloseAccountCommission_ist.xls")     '程式別不同
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

        'NO
        Dim wNo As String = ""
        If DNO.Text <> "" Then
            wNo = " and a.No like  '%" + Trim(DNO.Text) + "%'"
        Else
            wNo = ""
        End If



        'NO
        Dim wName As String = ""
        If DName.Text <> "" Then
            wName = " and EmpName2  like '%" + DName.Text + "%'"
        Else
            wName = ""
        End If




        '狀態
        Dim wSts As String = ""
        If DSTS.SelectedValue <> "" Then
            wSts = " and b.Sts in (" + DSTS.Text + ")"
        Else
            wSts = ""
        End If




        Dim Division As String = ""
        If DDivision1.SelectedValue <> "" Then
            Division = " and Division2 ='" + DDivision1.SelectedValue + "'"
        End If



        Dim CostDivision As String = ""
        If DDivision2.SelectedValue <> "" Then
            CostDivision = " and Division3 ='" + DDivision2.SelectedValue + "'"
        End If



        '出差
        '開始日期
        Dim Sdate As Date
        Dim Sdate1 As String = ""
        If DSDate.Text <> "" Then
            Sdate = DSDate.Text
            Sdate1 = Format(Sdate, "yyyy/MM/dd")
            Sdate1 = " and  Convert(VARCHAR(10), DATE, 111) >= '" + Sdate1 + "'"
        End If

        '結束日期
        Dim Edate As Date
        Dim Edate1 As String = ""

        If DEDate.Text <> "" Then
            Edate = DEDate.Text
            Edate1 = Format(Edate, "yyyy/MM/dd")
            Edate1 = " and  Convert(VARCHAR(10),DATE, 111) <= '" + Edate1 + "'"
        End If




        '預定
        '開始日期
        Dim ASdate As Date
        Dim ASdate1 As String = ""
        If DASDate.Text <> "" Then
            ASdate = DASDate.Text
            ASdate1 = Format(ASdate, "yyyy/MM/dd")
            ASdate1 = " and  Convert(VARCHAR(10), ASDATE, 111) >= '" + ASdate1 + "'"
        End If

        '結束日期
        Dim AEdate As Date
        Dim AEdate1 As String = ""

        If DAEDate.Text <> "" Then
            AEdate = DAEDate.Text
            AEdate1 = Format(AEdate, "yyyy/MM/dd")
            AEdate1 = " and  Convert(VARCHAR(10),AEDATE, 111) <= '" + AEdate1 + "'"
        End If




        Dim i As Integer = 0
        Dim SQL As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "出差者 "
        DataGrid1.Columns.Item(3).HeaderText = "出差者部門"
        DataGrid1.Columns.Item(4).HeaderText = "費用歸屬部門"
        DataGrid1.Columns.Item(5).HeaderText = "申請日期"
        DataGrid1.Columns.Item(6).HeaderText = "實際日程(起)"
        DataGrid1.Columns.Item(7).HeaderText = "實際日程(迄)"
        DataGrid1.Columns.Item(8).HeaderText = "目的"
        DataGrid1.Columns.Item(9).HeaderText = "地區/訪問"


        SQL = "SELECT　a.no as Field1,case when b.sts='0' then '核定中' when b.sts ='1' then '完成' else '取消' end as Field2,"
        SQL = SQL + " EmpName2 as Field3, Division2 as Field4,Division3 as Field5, convert(char(10),Date,111) as Field6, convert(char(10),aSDate,111)  as Field7, convert(char(10),aEDate,111)  as Field8,Object as Field9,Location as Field10,  "
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
        SQL = SQL + "where formno = 3115"
        SQL = SQL + "  and step =1) a,f_CloseAccountSheet b"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'  "
        SQL = SQL + wNo + wName + Sdate1 + Edate1 + ASdate1 + AEdate1 + wSts
        ' + Sdate1 + Edate1 + +ASdate1 + AEdate1 + wSts
        SQL = SQL + " order by  SDate desc "
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





        '出差部門



        SQL = " Select distinct depid,depname as Data  From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        SQL = SQL & "  where a.empname =b.username"
        SQL = SQL & "   And Active = '1' order by depid  "




        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)

        DDivision1.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DDivision1.Items.Add(ListItem1)

        Next
        dtReferp.Clear()


        '費用部門


        SQL = " Select distinct costid,costname as Data  From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
        SQL = SQL & "  where a.empname =b.username"
        SQL = SQL & "   And Active = '1' order by costid   "




        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)

        DDivision2.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DDivision2.Items.Add(ListItem1)

        Next
        dtReferp1.Clear()








        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"

    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class

