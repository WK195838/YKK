Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class BusinessTripCheckList
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
    Dim wStep, wDelay, wDep, wType As String
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
        wStep = Request.QueryString("pStep")
        wDelay = Request.QueryString("pDelay")
        wDep = Request.QueryString("pDep")
        wType = Request.QueryString("pType")


        BSDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"

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


    Sub DataList()


        'NO
        Dim wNo As String = ""
        If DNO.Text <> "" Then
            wNo = " and No like  '%" + Trim(DNO.Text) + "%'"
        Else
            wNo = ""
        End If


        'NO
        Dim wName As String = ""
        If DName.Text <> "" Then
            wName = " and EmpName2  like '%" + Trim(DName.Text) + "%'"
        Else
            wName = ""
        End If



        '開始日期
        Dim ASdate As Date
        Dim ASdate1 As String = ""
        If DASDate.Text <> "" Then
            ASdate = DASDate.Text
            ASdate1 = Format(ASdate, "yyyy/MM/dd")
            ASdate1 = " and  Convert(VARCHAR(10), SDATE, 111) >= '" + ASdate1 + "'"
        End If

        '結束日期
        Dim AEdate As Date
        Dim AEdate1 As String = ""

        If DAEDate.Text <> "" Then
            AEdate = DAEDate.Text
            AEdate1 = Format(AEdate, "yyyy/MM/dd")
            AEdate1 = " and  Convert(VARCHAR(10),EDATE, 111) <= '" + AEdate1 + "'"
        End If


        '狀態
        Dim wSts As String = ""
        If DSTS.SelectedValue <> "" Then
            If DSTS.Text = "清算中" Then
                wSts = " and Sts =0"
            ElseIf DSTS.Text = "清算完畢" Then
                wSts = " and Sts =1"
            Else
                wSts = " and Sts =9"
            End If

        Else
            wSts = ""
        End If


        'NO
        Dim wObject As String = ""
        If DObject.Text <> "" Then
            wObject = " and Object  like '%" + Trim(DObject.Text) + "%'"
        Else
            wObject = ""
        End If

        'NO
        Dim wLocation As String = ""
        If DLocation.Text <> "" Then
            wLocation = " and Location  like '%" + Trim(DLocation.Text) + "%'"
        Else
            wLocation = ""
        End If






        Dim sql As String

        sql = "  select No as '出差申請單號',empname2 as '出差者',sdate as '預定日程(起)',edate as '預定日程(迄)',object as '目的',location as '地區/訪問',"
        sql = sql & " ClostAccNo as '清算申請單號',  case when sts =0 then '清算中' when sts =1 then '清算完畢' else '未清算' end 狀態,cformsno,bformsno  from ( "
        sql = sql & " select * from ("
        sql = sql & " select a.*,isnull(b.no,'')ClostAccNo,isnull(b.sts,'9')sts,isnull(b.formsno,'')Cformsno,a.formsno as bformsno  from ("
        sql = sql & " select formsno,no,empname2,convert(char(10),sdate,111)sdate,convert(char(10),edate,111)edate,object,location,CloseAccNo  from f_BusinessTripSheet  "
        sql = sql & " where(sts = 1) )a left join  f_CloseAccountSheet b  on a.CloseAccNo =b.No  "
        sql = sql & " )a where 1=1" + wNo + wName + ASdate1 + AEdate1 + wObject + wLocation + wSts
        sql = sql & " )a "





        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        Dim BFormsno, CFormsno As String
        Dim h1, h2 As New HyperLink

        If e.Row.RowType = DataControlRowType.DataRow Then
            BFormsno = e.Row.Cells(9).Text
            h1.Text = e.Row.Cells(0).Text
            ' 連結到待處理LIST

            h1.NavigateUrl = "BusinessTripSheet_02.aspx?&pFormno=003114&pFormsno=" + BFormsno
            h1.Target = "_blank"
            e.Row.Cells(0).Text = ""
            e.Row.Cells(0).Controls.Add(h1)

            If e.Row.Cells(8).Text <> "0" Then
                CFormsno = e.Row.Cells(8).Text
                h2.Text = e.Row.Cells(6).Text
                '連結到待處理LIST()

                h2.NavigateUrl = "CloseAccountSheet_02.aspx?&pFormno=003115&pFormsno=" + CFormsno
                h2.Target = "_blank"
                e.Row.Cells(6).Text = ""
                e.Row.Cells(6).Controls.Add(h2)

            End If


        End If

        e.Row.Cells(8).Visible = False
        e.Row.Cells(9).Visible = False


    End Sub

    Protected Sub BExcel_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=BusinessTripCheckList.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView1.AllowPaging = wAllowPaging        '程式別不同


    End Sub

 
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataList()
    End Sub
End Class

