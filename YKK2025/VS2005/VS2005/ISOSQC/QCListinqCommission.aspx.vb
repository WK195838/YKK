Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class QCListinqCommission
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

    'IRW-EDX LINK
    'ADD-START BY JOY 230926
    Dim pIRW As Boolean
    Dim pSize As String
    Dim pSlider As String
    Dim pPuller As String
    'ADD-EBD BY JOY 230926
    Dim pWIP As String

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
        BQCDate.Attributes("onclick") = "calendarPicker('Form1.DQCDate');"

        'IRW-EDX LINK
        'ADD-START BY JOY 230926
        If Not IsPostBack Then
            pIRW = Request.QueryString("pIRW")
            pSize = Request.QueryString("pSize")
            pSlider = Request.QueryString("pSlider")
            pPuller = Request.QueryString("pPuller")
            If pIRW = "1" Then
                pIRW = "0"
                DataGrid1.CurrentPageIndex = 0
                DSPEC.Text = pSlider
                DataListIRW()
            End If

            '從ISIP EDX=WIP進來，查詢實際狀態
            pWIP = Request.QueryString("pWIP")
            If pWIP = "1" Then
                DSPEC.Text = pSlider
                Go_Click(Go, New System.EventArgs)
            End If
        End If

        'ADD-EBD BY JOY 230926
    End Sub

    'IRW-EDX LINK
    'ADD-START BY JOY 230926
    Sub DataListIRW()
        Dim SQL As String = ""
        '
        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "依賴日 "
        DataGrid1.Columns.Item(3).HeaderText = "申請者"
        DataGrid1.Columns.Item(4).HeaderText = "客戶"
        DataGrid1.Columns.Item(5).HeaderText = "BUYER"
        DataGrid1.Columns.Item(6).HeaderText = "品管受理NO"
        DataGrid1.Columns.Item(7).HeaderText = "分析完了日"
        DataGrid1.Columns.Item(8).HeaderText = "型別組"
        DataGrid1.Columns.Item(9).HeaderText = "TestResult"
        DataGrid1.Columns.Item(10).HeaderText = "覆判結果"
        '
        SQL = "SELECT　distinct a.no as Field1,case when b.sts='0' then '核定中' when b.sts ='1' then '完成' else '取消' end as Field2,"
        SQL = SQL + " convert(char(10),ADate,111)  as Field3, Name as Field4,Customer as Field5,b.Buyer as Field6,QCNO as Field7,"
        SQL = SQL + " case when qcdate ='1900-01-01 00:00:00.000' then '' else convert(char(10),QCDate,111) end as Field8,'('+Supplier+') '+Size+' '+Family+' '+Body+' '+Puller+' '+Color+' '+Finish as Field9,"
        SQL = SQL + " result1 Field10,result2 Field11 ,"
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
        SQL = SQL + "where formno = 8002"
        SQL = SQL + "  and step =1) a,f_QASheet b ,F_QASheetdt c "
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and b.no =a.no and b.no =c.no "
        '
        '一般SLIDER CODE (SPD,QC)
        SQL = SQL & "And ( "
        SQL = SQL & "( "
        SQL = SQL & "      Puller + Color Like '%" & Mid(pSlider, 3, 999) & "%' "
        SQL = SQL & "   Or Puller + Color Like '%" & Mid(pSlider, 4, 999) & "%' "
        SQL = SQL & "   Or Puller + Color Like '%" & Mid(pSlider, 5, 999) & "%' "
        SQL = SQL & "   Or Puller + Color Like '%" & Mid(pSlider, 6, 999) & "%' "
        SQL = SQL & ") "
        'SLIDER CODE縮寫(PL-)
        If pPuller <> "" Then
            'QC EDX
            'IRW: DA7BA PL-BS054
            'EDX : PL-BS054 A
            SQL = SQL & "Or ( "
            SQL = SQL & "      Color + Puller Like '%" & Mid(pSlider + pPuller, 3, 999) & "%' "
            SQL = SQL & "   Or Color + Puller Like '%" & Mid(pSlider + pPuller, 4, 999) & "%' "
            SQL = SQL & "   Or Color + Puller Like '%" & Mid(pSlider + pPuller, 5, 999) & "%' "
            SQL = SQL & "   Or Color + Puller Like '%" & Mid(pSlider + pPuller, 6, 999) & "%' "
            SQL = SQL & ") "
            '(特殊) PULLER CODE分離 (PL-LLM)
            'IRW: DA7B74A  PL-LLM
            'EDX : PL-LLM74 A
            SQL = SQL & "Or ( "
            SQL = SQL & "      Puller + Color Like '%" & pPuller + Mid(pSlider, 3, 999) & "%' "
            SQL = SQL & "   Or Puller + Color Like '%" & pPuller + Mid(pSlider, 4, 999) & "%' "
            SQL = SQL & "   Or Puller + Color Like '%" & pPuller + Mid(pSlider, 5, 999) & "%' "
            SQL = SQL & "   Or Puller + Color Like '%" & pPuller + Mid(pSlider, 6, 999) & "%' "
            SQL = SQL & ") "
            'PRICE
            SQL = SQL & "Or ( "
            SQL = SQL & "      Puller Like '%" & pSlider & "%' "
            SQL = SQL & ") "
        End If
        SQL = SQL & ") "
        '
        SQL = SQL + " order by  a.no desc "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DataGrid1.DataSource = DBAdapter1
        DataGrid1.DataBind()
    End Sub
    'ADD-EBD BY JOY 230926


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

        Response.AppendHeader("Content-Disposition", "attachment;filename=QAEDXResult_List.xls")     '程式別不同
        '  Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
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


        Dim Spec As String = ""
        If DSPEC.Text <> "" Then
            Spec = " and size+family+body+puller+color+finish like '%" + DSPEC.Text.Replace(" ", "") + "%'"
        End If

        Dim qcno As String = ""
        If DQCNo.Text <> "" Then
            qcno = " and qcno  like '%" + DQCNo.Text.Replace(" ", "") + "%'"
        End If


 
        '依賴日期
        Dim Sdate As Date
        Dim Sdate1 As String = ""
        If DADate.Text <> "" Then
            Sdate = DADate.Text
            Sdate1 = Format(Sdate, "yyyy/MM/dd")
            Sdate1 = " and  Convert(VARCHAR(10), ADATE, 111) >= '" + Sdate1 + "'"
        End If

        '客訴完了日
        Dim Edate As Date
        Dim Edate1 As String = ""

        If DQCDate.Text <> "" Then
            Edate = DQCDate.Text
            Edate1 = Format(Edate, "yyyy/MM/dd")
            Edate1 = " and  Convert(VARCHAR(10),QCDate, 111) <= '" + Edate1 + "'"
        End If



 




        Dim i As Integer = 0
        Dim SQL As String = ""

        DataGrid1.Columns.Item(0).HeaderText = "NO"
        DataGrid1.Columns.Item(1).HeaderText = "狀態"
        DataGrid1.Columns.Item(2).HeaderText = "依賴日 "
        DataGrid1.Columns.Item(3).HeaderText = "申請者"
        DataGrid1.Columns.Item(4).HeaderText = "客戶"
        DataGrid1.Columns.Item(5).HeaderText = "BUYER"
        DataGrid1.Columns.Item(6).HeaderText = "品管受理NO"
        DataGrid1.Columns.Item(7).HeaderText = "分析完了日"
        DataGrid1.Columns.Item(8).HeaderText = "型別組"
        DataGrid1.Columns.Item(9).HeaderText = "TestResult"
        DataGrid1.Columns.Item(10).HeaderText = "覆判結果"


        SQL = "SELECT　distinct a.no as Field1,case when b.sts='0' then '核定中' when b.sts ='1' then '完成' else '取消' end as Field2,"
        SQL = SQL + " convert(char(10),ADate,111)  as Field3, Name as Field4,Customer as Field5,b.Buyer as Field6,QCNO as Field7,"
        SQL = SQL + " case when qcdate ='1900-01-01 00:00:00.000' then '' else convert(char(10),QCDate,111) end as Field8,'('+Supplier+') '+Size+' '+Family+' '+Body+' '+Puller+' '+Color+' '+Finish as Field9,"
        SQL = SQL + " result1 Field10,result2 Field11 ,"
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
        SQL = SQL + "where formno = 8002"
        SQL = SQL + "  and step =1) a,f_QASheet b ,F_QASheetdt c "
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and b.no =a.no and b.no =c.no "
        SQL = SQL + wNo + wName + Sdate1 + Edate1 + wSts + Buyer + Customer + Spec + qcno
        ' + Sdate1 + Edate1 + +ASdate1 + AEdate1 + wSts
        SQL = SQL + " order by  a.no desc "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DataGrid1.DataSource = DBAdapter1
        DataGrid1.DataBind()

    End Sub

    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        pIRW = "0"
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

