Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class ScheduleBC
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
    Dim ItemCount As Integer = 6  '預先定義欄位數量 





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
        SetMainMenu()         '設定主畫面
        If Not Me.IsPostBack Then
            DataList()
            DSTS.SelectedValue = "未清算"
            BusinessTripCheckList()
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
        Dim sql As String
 
        sql = "select * from ( "
        sql = sql + "select a.step,a.stepname,isnull(allcase,0)allcase, case when allcase is null or allcase =0 then '' else '...' end as allcaseURL, "
        sql = sql + "isnull(delaycase,0)delaycase, case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL  from ( "
        sql = sql + "select  step,substring(stepname,6,len(stepname)-1)stepname from M_leadtime where formno ='003114'  )a left join "
        sql = sql + "( select step,sum(allcase)allcase,sum(delaycase)delaycase from( select formno,step,substring(stepname,11,len(stepname)-1)stepname,1 as allcase, "
        sql = sql + "case when delaysts ='正常' then  0 else 1 end as delaycase,formsno  from   V_WaitHandle_01  where formno = '003114' and flowtype ='1' and active ='1'  "
        sql = sql + "and seqno=1  "
        sql = sql + ")a,( select formno,formsno  from F_BusinessTripSheet "
        sql = sql + "where sts =0  )b  where a.formsno =b.formsno and b.formno = '003114' group by step )b on a.step =b.step )a where step in (10,20,30,40,50,500  ) order by step"

  

        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()



        sql = "select * from ( "
        sql = sql + "select a.step,a.stepname,isnull(allcase,0)allcase, case when allcase is null or allcase =0 then '' else '...' end as allcaseURL, "
        sql = sql + "isnull(delaycase,0)delaycase, case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL  from ( "
        sql = sql + "select  step,substring(stepname,6,len(stepname)-1)stepname from M_leadtime where formno ='003115'  )a left join "
        sql = sql + "( select step,sum(allcase)allcase,sum(delaycase)delaycase from( select formno,step,substring(stepname,11,len(stepname)-1)stepname,1 as allcase, "
        sql = sql + "case when delaysts ='正常' then  0 else 1 end as delaycase,formsno  from   V_WaitHandle_01  where formno = '003115' and flowtype ='1' and active ='1'  "
        sql = sql + "and seqno=1  "
        sql = sql + ")a,( select formno,formsno  from F_CloseAccountSheet "
        sql = sql + "where sts =0  )b  where a.formsno =b.formsno and b.formno = '003115' group by step )b on a.step =b.step )a where step in (10,20,30,40,50,60,500 ) order by step"



        Dim dtData1 As DataTable = uDataBase.GetDataTable(sql)
        GridView2.DataSource = dtData1
        GridView2.DataBind()


    End Sub




    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If


        If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then


            Dim h1 As New HyperLink
            Dim h2 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(5).Text

            If e.Row.Cells(0).Text <> "合計" Then
                If Trim(e.Row.Cells(1).Text) <> "0" Then
                    h1.Target = "_blank"
                    h1.Text = e.Row.Cells(2).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h1.NavigateUrl = "ScheduleAllBC.aspx?&pFormNo=003114&pStep=" + wStep
                    '+ "&pUserID=" + Request.QueryString("pUserID")

                    ' e.Row.Cells(3).Text = ""
                    e.Row.Cells(2).Controls.Add(h1)

                End If

                If Trim(e.Row.Cells(3).Text) <> "0" Then
                    h2.Target = "_blank"
                    h2.Text = e.Row.Cells(4).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h2.NavigateUrl = "ScheduleAllBC.aspx?&pFormNo=003114&pStep=" + wStep
                    '+ "&pUserID=" + Request.QueryString("pUserID")

                    ' e.Row.Cells(4).Text = ""
                    e.Row.Cells(4).Controls.Add(h2)

                End If


            End If

            e.Row.Cells(5).Visible = False


            '連結
        End If


    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0

        For Each gvr As GridViewRow In GridView1.Rows

            intTotal += gvr.Cells(1).Text
            intTotal1 += gvr.Cells(3).Text

            gvr.Cells(1).Text = FormatNumber(gvr.Cells(1).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(1).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            GridView1.FooterRow.Cells(1).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView1.FooterRow.Cells(1).HorizontalAlign = HorizontalAlign.Center

            gvr.Cells(3).Text = FormatNumber(gvr.Cells(3).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            GridView1.FooterRow.Cells(3).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            GridView1.FooterRow.Cells(3).HorizontalAlign = HorizontalAlign.Center


        Next



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定菜單各程式 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim xStep As Integer = 1
        Dim SQL As String = ""


        LFun01.Enabled = True
        LFun02.Enabled = True


        LFun01.NavigateUrl = "BusinessTripSheet_01.aspx?pFormNo=003114" & _
                                                                            "&pFormSno=0" & _
                                                                            "&pStep=" & xStep & _
                                                                            "&pSeqNo=0" & _
                                                                            "&pUserID=" & Request.QueryString("pUserID") & _
                                                                            "&pApplyID=" & Request.QueryString("pUserID")
        LFun02.NavigateUrl = "BusinessTripListinqCommission.aspx?pFormNo=003114" & _
                                                         "&pUserID=" & Request.QueryString("pUserID")


        LFun01.Target = "_blank"

        LFun02.Target = "_blank"


        LFun03.Enabled = True
        LFun04.Enabled = True


        LFun03.NavigateUrl = "CloseAccountSheet_01.aspx?pFormNo=003115" & _
                                                                            "&pFormSno=0" & _
                                                                            "&pStep=" & xStep & _
                                                                            "&pSeqNo=0" & _
                                                                            "&pUserID=" & Request.QueryString("pUserID") & _
                                                                            "&pApplyID=" & Request.QueryString("pUserID")
        LFun04.NavigateUrl = "CloseAccountinqCommission.aspx?pFormNo=003115" & _
                                                         "&pUserID=" & Request.QueryString("pUserID")


        LFun03.Target = "_blank"
        LFun04.Target = "_blank"

  
        '  LFun02.Target = "_blank"

    End Sub


    Protected Sub LFun01_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun02_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun02.Init
        LFun02.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun02.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub


    Protected Sub LFun03_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            ' 建立自訂的標題 

            Dim gv As GridView = DirectCast(sender, GridView)
            Dim gvRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            Dim tc1 As New TableCell()
            tc1.Text = "工程"
            tc1.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1)

            Dim tc2 As New TableCell()
            tc2.Text = "待處理件數"
            tc2.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc2)

            Dim tc3 As New TableCell()
            tc3.Text = "待處理連結"
            tc3.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc3)


            Dim tc4 As New TableCell()
            tc4.Text = "延遲件數"
            tc4.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc4)

            Dim tc5 As New TableCell()
            tc5.Text = "延遲連結"
            tc5.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc5)



            ' Dim tc6 As New TableCell()
            'tc6.Text = "STEP"
            'tc6.HorizontalAlign = HorizontalAlign.Center
            'gvRow.Cells.Add(tc6)

            e.Row.Cells.Clear()
            gv.Controls(0).Controls.AddAt(0, gvRow)

        End If

    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If


        If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then


            Dim h1 As New HyperLink
            Dim h2 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(5).Text

            If e.Row.Cells(0).Text <> "合計" Then
                If Trim(e.Row.Cells(1).Text) <> "0" Then
                    h1.Target = "_blank"
                    h1.Text = e.Row.Cells(2).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h1.NavigateUrl = "ScheduleAllBC.aspx?&pFormNo=003115&pStep=" + wStep
                    ' + "&pUserID=" + Request.QueryString("pUserID")

                    ' e.Row.Cells(3).Text = ""
                    e.Row.Cells(2).Controls.Add(h1)

                End If

                If Trim(e.Row.Cells(3).Text) <> "0" Then
                    h2.Target = "_blank"
                    h2.Text = e.Row.Cells(4).Text
                    ' 連結到待處理LIST
                    ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                    'h1.NavigateUrl = "FundingScheduleAll.aspx?&pStep=" + wStep
                    h2.NavigateUrl = "ScheduleAllBC.aspx?&pFormNo=003115&pStep=" + wStep
                    '+ "&pUserID=" + Request.QueryString("pUserID")

                    ' e.Row.Cells(4).Text = ""
                    e.Row.Cells(4).Controls.Add(h2)

                End If


            End If

            e.Row.Cells(5).Visible = False


            '連結
        End If

    End Sub

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            ' 建立自訂的標題 

            Dim gv As GridView = DirectCast(sender, GridView)
            Dim gvRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            Dim tc1 As New TableCell()
            tc1.Text = "工程"
            tc1.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc1)

            Dim tc2 As New TableCell()
            tc2.Text = "待處理件數"
            tc2.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc2)

            Dim tc3 As New TableCell()
            tc3.Text = "待處理連結"
            tc3.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc3)


            Dim tc4 As New TableCell()
            tc4.Text = "延遲件數"
            tc4.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc4)

            Dim tc5 As New TableCell()
            tc5.Text = "延遲連結"
            tc5.HorizontalAlign = HorizontalAlign.Center
            gvRow.Cells.Add(tc5)



            ' Dim tc6 As New TableCell()
            'tc6.Text = "STEP"
            'tc6.HorizontalAlign = HorizontalAlign.Center
            'gvRow.Cells.Add(tc6)

            e.Row.Cells.Clear()
            gv.Controls(0).Controls.AddAt(0, gvRow)

        End If

    End Sub

    Protected Sub GridView2_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0

        For Each gvr As GridViewRow In GridView2.Rows

            intTotal += gvr.Cells(1).Text
            intTotal1 += gvr.Cells(3).Text

            gvr.Cells(1).Text = FormatNumber(gvr.Cells(1).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(1).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            GridView2.FooterRow.Cells(1).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView2.FooterRow.Cells(1).HorizontalAlign = HorizontalAlign.Center

            gvr.Cells(3).Text = FormatNumber(gvr.Cells(3).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            GridView2.FooterRow.Cells(3).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            GridView2.FooterRow.Cells(3).HorizontalAlign = HorizontalAlign.Center


        Next

    End Sub


    Sub BusinessTripCheckList()


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
        GridView3.DataSource = dtData
        GridView3.DataBind()
    End Sub

    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        BusinessTripCheckList()
    End Sub

    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView3.AllowPaging = False   '程式別不同
        BusinessTripCheckList()                  '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=BusinessTripCheckList.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView3.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView3.AllowPaging = wAllowPaging        '程式別不同

    End Sub

    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
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
End Class

