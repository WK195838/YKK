Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class QCModListinqCommission
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

        BADate.Attributes("onclick") = "calendarPicker('Form1.DADate');"




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

        Response.AppendHeader("Content-Disposition", "attachment;filename=BusinessTripCommission_ist.xls")     '程式別不同
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
            wName = " and Name  like '%" + DName.Text + "%'"
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


        Dim Customer As String = ""
        If DCustomer.Text <> "" Then
            Customer = " and (Customer like '%" + DCustomer.Text + "%' or CustomerCode like '%" + DCustomer.Text + "%' )  "
        End If



        Dim Buyer As String = ""
        If DBuyer.Text <> "" Then
            Buyer = " and (b.Buyer like '%" + DBuyer.Text + "%' or b.BuyerCode like '%" + DBuyer.Text + "%' ) "
        End If



        '依賴日期
        Dim Sdate As Date
        Dim Sdate1 As String = ""
        If DADate.Text <> "" Then
            Sdate = DADate.Text
            Sdate1 = Format(Sdate, "yyyy/MM/dd")
            Sdate1 = " and  Convert(VARCHAR(10), DATE, 111) = '" + Sdate1 + "'"
        End If


        Dim ModNo As String = ""

        If DModNo.Text <> "" Then
            ModNo = " and ModNo  like '%" + DModNo.Text + "%'"
        End If

        Dim Puller As String = ""
        If DPuller.Text <> "" Then
            Puller = " And Size+Family+Body+Puller+Color+Finish Like '%" + DPuller.Text.Replace(" ", "") + "%'"
        End If



        Dim i As Integer = 0
        Dim SQL As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "申請日"
        DataGrid1.Columns.Item(3).HeaderText = "申請者"
        DataGrid1.Columns.Item(4).HeaderText = "原始單號"
        DataGrid1.Columns.Item(5).HeaderText = "修改理由"
        DataGrid1.Columns.Item(6).HeaderText = "修改內容"
        DataGrid1.Columns.Item(7).HeaderText = "型別組"


        SQL = "SELECT DISTINCT　a.no as Field1,case when b.sts='0' then '核定中' when b.sts ='1' then '完成' else '取消' end as Field2,"
        SQL = SQL + " convert(char(10),Date,111)  as Field3, Name as Field4,ModNO as Field5,ModReason as Field6,ModContent as Field7,'('+Supplier+') '+Size+' '+Family+' '+Body+' '+Puller+' '+Color+' '+Finish as Field8,"
        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + a.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + a.ApplyID "
        SQL = SQL + " As OPURL,  "
        SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate "
        SQL = SQL + " from (select * from   V_WaitHandle_01 "
        SQL = SQL + " where formno = '008003' "
        SQL = SQL + "  and step =1) a,f_QAModSheet b,f_QAModSheetDT c "
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and b.no = a.no and b.no =c.no "
        SQL = SQL + wNo + wName + Sdate1 + ModNo + wSts + Buyer + Customer + Puller
        ' + Sdate1 + Edate1 + +ASdate1 + AEdate1 + wSts
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

End Class

