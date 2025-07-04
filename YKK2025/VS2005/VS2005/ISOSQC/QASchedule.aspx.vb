Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class QASchedule
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
        SetMainMenu()
        If Not Me.IsPostBack Then
            DataList()

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

        '  BQACHECK.Attributes.Add("onClick", "RunExcelCHECK()") ' 查詢

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

        'sql = " select a.step,substring(b.stepname,9,len(b.stepname)-1)stepname,sum(allcase)allcase,allcaseURL,sum(delayCase)delayCase,delaycaseURL,a.Leadtime  from ("
        'sql = sql + " select a,formno,a.step,substring(stepname,9,len(stepname)-1)stepname,"
        'sql = sql + " case when allcase is null then 0 else allcase end as allcase, case when allcase is null or allcase =0"
        'sql = sql + " then '' else '...' end as allcaseURL, case when delaycase is null then  0 else delaycase end as delaycase,"
        'sql = sql + " case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL,Leadtime from ( "
        'sql = sql + " select a.*,hours,Hours/8/60 as Leadtime from("
        'sql = sql + " select distinct formno,step,stepname  from M_flow a where formno = 003109 and step not in (1,999,40,15,35,31,41) "
        'sql = sql + " )a,M_leadtime b"
        'sql = sql + " where(a.step = b.step And a.formno = b.formno)"
        'sql = sql + " )a left join ("
        'sql = sql + " select step,stepname as stepname1,sum(allcase)allcase,sum(delaycase)delaycase from ("
        'sql = sql + " select step,substring(stepname,9,len(stepname)-1)stepname,1 as allcase,"
        'sql = sql + " case when delaysts ='正常' then  0 else 1 end as delaycase  from   V_WaitHandle_01 "
        'sql = sql + " where formno = '003109' and flowtype ='1' and active ='1'  and step <> '40' and seqno=1"
        'sql = sql + " )a group by step,stepname "
        'sql = sql + " ) b on a.step =b.step )a "
        'sql = sql + " ,m_flow b"
        'sql = sql + " where  a.step = b.step And Action = 0 And a.formno = b.formno)"
        'sql = sql + " group by  a.step,a.allcaseURL,a.delaycaseURL,a.Leadtime,b.stepname"
        'sql = sql + " order by a.step"



        sql = "   select a.step,substring(b.stepname,9,len(b.stepname)-1)stepname,sum(allcase)allcase,allcaseURL,sum(delayCase)delayCase,delaycaseURL "
        sql = sql + "  from ("
        sql = sql + "  select a.formno,a.step,substring(stepname,9,len(stepname)-1)stepname, "
        sql = sql + "  case when allcase is null then 0 else allcase end as allcase, "
        sql = sql + " case when allcase is null then '' else '...' end as allcaseURL, "
        sql = sql + "  case when delaycase  is null then  0 else delaycase end as delaycase, "
        sql = sql + "  case when delaycase is null or delaycase =0  then ''  else '...'  end as delaycaseURL,Leadtime from ("
        sql = sql + "  select a.*,hours,Hours/8/60 as Leadtime from( "
        sql = sql + "  select distinct formno,step,stepname  from M_flow a where formno = 008002 "
        sql = sql + "  and step not in (1,999,21)  )a,M_leadtime b "
        sql = sql + "  where(a.step = b.step And a.formno = b.formno) )a left join ( "
        sql = sql + "  select step,stepname as stepname1,sum(allcase)allcase,sum(delaycase)delaycase from"
        sql = sql + "  ( select step,substring(stepname,9,len(stepname)-1)stepname,1 as allcase, "
        sql = sql + "  case when delaysts ='正常' then  0 else 1 end as delaycase  from   V_WaitHandle_01"
        sql = sql + "  where formno = '008002' and flowtype ='1' and active ='1' "
        sql = sql + "  and seqno=1 )a group by step,stepname  ) b on a.step =b.step"
        sql = sql + "  )a,m_flow b"
        sql = sql + "  where(a.step = b.step And Action = 0 And a.formno = b.formno)"
        sql = sql + "  group by  a.step,a.allcaseURL,a.delaycaseURL,a.Leadtime,b.stepname"
        sql = sql + "  order by a.step"



        Dim dtData1 As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData1
        GridView1.DataBind()


        sql = " select a.step,substring(b.stepname,9,len(b.stepname)-1)stepname,sum(allcase)allcase,allcaseURL,sum(delayCase)delayCase,delaycaseURL "
        sql = sql + " from ("
        sql = sql + " select a.formno,a.step,substring(stepname,12,len(stepname)-1)stepname, "
        sql = sql + " case when allcase is null then 0 else allcase end as allcase, "
        sql = sql + " case when allcase is null then '' else '...' end as allcaseURL, "
        sql = sql + " case when delaycase  is null then  0 else delaycase end as delaycase, "
        sql = sql + " case when delaycase is null or delaycase =0  then ''  else '...'  end as delaycaseURL from ( "
        sql = sql + " select distinct formno,step,stepname  from M_flow a where formno = '008003'"
        sql = sql + " and step not in (1,11,12,500,999) )a  left join ( "
        sql = sql + " select step,stepname as stepname1,sum(allcase)allcase,sum(delaycase)delaycase from"
        sql = sql + " ( select step,substring(stepname,9,len(stepname)-1)stepname,1 as allcase, "
        sql = sql + " case when delaysts ='正常' then  0 else 1 end as delaycase  from   V_WaitHandle_01"
        sql = sql + " where formno = '008003' and flowtype ='1' and active ='1' "
        sql = sql + " and seqno=1 )a group by step,stepname  ) b on a.step =b.step"
        sql = sql + " )a,m_flow b"
        sql = sql + " where(a.step = b.step And Action = 0 And a.formno = b.formno)"
        sql = sql + " group by  a.step,a.allcaseURL,a.delaycaseURL,b.stepname"
        sql = sql + " order by a.step"



        Dim dtData2 As DataTable = uDataBase.GetDataTable(sql)
        GridView2.DataSource = dtData2
        GridView2.DataBind()



        GridView3.Columns.Item(0).HeaderText = "NO"
        GridView3.Columns.Item(1).HeaderText = "狀態"
        GridView3.Columns.Item(2).HeaderText = "依賴日 "
        GridView3.Columns.Item(3).HeaderText = "申請者"
        GridView3.Columns.Item(4).HeaderText = "客戶"
        GridView3.Columns.Item(5).HeaderText = "BUYER"
        GridView3.Columns.Item(6).HeaderText = "型別組"
        GridView3.Columns.Item(7).HeaderText = "TestResult"
        GridView3.Columns.Item(8).HeaderText = "待判(element)"



        sql = "SELECT　distinct a.no as Field1,case when b.sts='0' then '核定中' when b.sts ='1' then '完成' else '取消' end as Field2,"
        sql = sql + " convert(char(10),ADate,111)  as Field3, Name as Field4,Customer as Field5,b.Buyer as Field6,"
        sql = sql + " '('+Supplier+') '+Size+' '+Family+' '+Body+' '+Puller+' '+Color+' '+Finish as Field7,"
        'Modify-Start by Joy 250605
        sql = sql + " case when step=500 then 'QC退件等待申請者處理' else result1 end as Field8, "
        sql = sql + " Element Field9, "
        'sql = sql + " result1 Field8,Element Field9 ,"
        'Modify-End by
        sql = sql + " '....' as WorkFlow, ViewURL, "
        sql = sql + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        sql = sql + "'pFormNo='   + a.FormNo + "
        sql = sql + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        sql = sql + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        sql = sql + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        sql = sql + "'&pApplyID=' + a.ApplyID "
        sql = sql + " As OPURL,  "
        sql = sql + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate"
        'Modify-Start by Joy 250605
        sql = sql + " from (select * from   V_WaitHandle_01 "
        sql = sql + "where formno = 8002"
        sql = sql + "  and step in (1,500)) a,f_QASheet b ,F_QASheetdt c "
        sql = sql + " Where "
        sql = sql + "      ( b.sts in(0,1) and a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and b.no =a.no and b.no =c.no and c.result1 not in ('OK','')  and c.result2 <>'OK' ) "
        sql = sql + "   or "
        sql = sql + "      ( b.sts in(0) and a.formno=b.formno and a.formsno =b.formsno and Step  = '500' and b.no =a.no and b.no =c.no ) "
        '
        'sql = sql + " from (select * from   V_WaitHandle_01 "
        'sql = sql + "where formno = 8002"
        'sql = sql + "  and step =1) a,f_QASheet b ,F_QASheetdt c "
        'sql = sql + " Where  b.sts in(0,1) and a.formno=b.formno and a.formsno =b.formsno and Step  = '1' and b.no =a.no and b.no =c.no and c.result1 not in ('OK','')  and c.result2 <>'OK' "
        'Modify-End

        sql = sql + " order by  a.no desc "

        Dim dtData3 As DataTable = uDataBase.GetDataTable(sql)
        GridView3.DataSource = dtData3
        GridView3.DataBind()



    End Sub





    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If


        If e.Row.RowType = DataControlRowType.DataRow Then


            Dim h1 As New HyperLink
            Dim h2 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(6).Text
            If e.Row.Cells(2).Text <> "0" Then

                h1.Target = "_blank"
                h1.Text = e.Row.Cells(3).Text
                ' 連結到待處理LIST
                ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                h1.NavigateUrl = "QAScheduleAll.aspx?&pStep=" + wStep
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If



            If e.Row.Cells(4).Text <> "0" Then
                h2.Target = "_blank"
                h2.Text = e.Row.Cells(5).Text
                ' 連結到延遲
                'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                h2.NavigateUrl = "QAScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                e.Row.Cells(5).Text = ""
                e.Row.Cells(5).Controls.Add(h2)
            End If

        End If
        '連結

        e.Row.Cells(6).Visible = False

    End Sub


    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0
        For Each gvr As GridViewRow In GridView1.Rows


            intTotal += gvr.Cells(2).Text

            intTotal1 += gvr.Cells(4).Text

            gvr.Cells(2).Text = FormatNumber(gvr.Cells(2).Text, formatNu, TriState.True, TriState.False, TriState.True)
            '  gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            GridView1.FooterRow.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView1.FooterRow.Cells(2).HorizontalAlign = HorizontalAlign.Center

            gvr.Cells(4).Text = FormatNumber(gvr.Cells(4).Text, formatNu, TriState.True, TriState.False, TriState.True)
            ' gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            GridView1.FooterRow.Cells(4).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            GridView1.FooterRow.Cells(4).HorizontalAlign = HorizontalAlign.Center
        Next



    End Sub
 
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If


        If e.Row.RowType = DataControlRowType.DataRow Then


            Dim h1 As New HyperLink
            Dim h2 As New HyperLink

            Dim wStep As String
            wStep = e.Row.Cells(6).Text
            If e.Row.Cells(2).Text <> "0" Then

                h1.Target = "_blank"
                h1.Text = e.Row.Cells(3).Text
                ' 連結到待處理LIST
                ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                h1.NavigateUrl = "QAModScheduleAll.aspx?&pStep=" + wStep
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If



            If e.Row.Cells(4).Text <> "0" Then
                h2.Target = "_blank"
                h2.Text = e.Row.Cells(5).Text
                ' 連結到延遲
                'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                h2.NavigateUrl = "QAModScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                e.Row.Cells(5).Text = ""
                e.Row.Cells(5).Controls.Add(h2)
            End If

        End If
        '連結

        e.Row.Cells(6).Visible = False

    End Sub


    Protected Sub GridView2_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0
        For Each gvr As GridViewRow In GridView2.Rows


            intTotal += gvr.Cells(2).Text

            intTotal1 += gvr.Cells(4).Text

            gvr.Cells(2).Text = FormatNumber(gvr.Cells(2).Text, formatNu, TriState.True, TriState.False, TriState.True)
            '  gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            GridView2.FooterRow.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView2.FooterRow.Cells(2).HorizontalAlign = HorizontalAlign.Center

            gvr.Cells(4).Text = FormatNumber(gvr.Cells(4).Text, formatNu, TriState.True, TriState.False, TriState.True)
            ' gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            GridView2.FooterRow.Cells(4).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            GridView2.FooterRow.Cells(4).HorizontalAlign = HorizontalAlign.Center
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
        LFun01.NavigateUrl = "QASheet_01.aspx?pFormNo=008002" & _
                                                          "&pFormSno=0" & _
                                                          "&pStep=" & xStep & _
                                                          "&pSeqNo=0" & _
                                                          "&pUserID=" & Request.QueryString("pUserID") & _
                                                          "&pApplyID=" & Request.QueryString("pUserID")
        LFun01.Target = "_blank"

        LFun02.Enabled = True
        'IRW-EDX LINK
        'MOD-START BY JOY 230926
        'LFun02.NavigateUrl = "QCListinqCommission.aspx"
        LFun02.NavigateUrl = "QCListinqCommission.aspx" & _
                              "?pIRW=0" & _
                              "&pSize=" & _
                              "&pSlider=" & _
                              "&pPuller=" 
        'MOD-EBD BY JOY 230926
        LFun02.Target = "_blank"


        LFun03.Enabled = True
        LFun03.NavigateUrl = "QAResultList.aspx"

        LFun03.Target = "_blank"


        LFun04.Enabled = True
        LFun04.NavigateUrl = "QCModListinqCommission.aspx"

        LFun04.Target = "_blank"


        LFun05.Enabled = True
        LFun05.NavigateUrl = "QAModSheet_01.aspx?pFormNo=008003" & _
                                                       "&pFormSno=0" & _
                                                       "&pStep=" & xStep & _
                                                       "&pSeqNo=0" & _
                                                       "&pUserID=" & Request.QueryString("pUserID") & _
                                                       "&pApplyID=" & Request.QueryString("pUserID")
        LFun05.Target = "_blank"

        Dim i As Integer
        Dim File(6) As String

        '找檔案路徑
        SQL = " select substring(data,3,len(data)-1)Data from  M_referp"
        SQL = SQL + " where cat = 8002  "
        SQL = SQL + " and dkey = 'file' and left(data,1) <>'7' "

        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter3.Rows.Count > 0 Then

            For i = 0 To DBAdapter3.Rows.Count - 1

                File(i + 1) = DBAdapter3.Rows(i).Item("Data")

            Next


        End If


        LEDX1.Enabled = True
        LEDX1.NavigateUrl = File(1)
        LEDX1.Target = "_blank"

        LEDX2.Enabled = True
        LEDX2.NavigateUrl = File(2)
        LEDX2.Target = "_blank"

        LEDX3.Enabled = True
        LEDX3.NavigateUrl = File(3)
        LEDX3.Target = "_blank"

        LEDX4.Enabled = True

        LEDX4.NavigateUrl = File(4)
        LEDX4.Target = "_blank"

        LEDX5.Enabled = True
        LEDX5.NavigateUrl = File(5)
        LEDX5.Target = "_blank"

        LEDX6.Enabled = True
        LEDX6.NavigateUrl = File(6)
        LEDX6.Target = "_blank"





    End Sub


    Protected Sub LFun01_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun02_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun02.Init
        LFun02.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun02.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

 
    Protected Sub LFun05_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun05.Init
        LFun05.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun05.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun04_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun04.Init
        LFun04.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun04.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun03_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun03.Init
        LFun03.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun03.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LEDX1_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LEDX1.Init
        LEDX1.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LEDX1.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LEDX2_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LEDX2.Init
        LEDX2.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LEDX2.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LEDX3_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LEDX3.Init
        LEDX3.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LEDX3.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LEDX4_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LEDX4.Init
        LEDX4.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LEDX4.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LEDX5_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LEDX5.Init
        LEDX5.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LEDX5.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LEDX6_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LEDX6.Init
        LEDX6.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LEDX6.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub GridView3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView3.SelectedIndexChanged

    End Sub

    Protected Sub GridView3_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView3.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0
        For Each gvr As GridViewRow In GridView3.Rows


            intTotal = intTotal + 1

            'intTotal1 += gvr.Cells(4).Text

            'gvr.Cells(2).Text = FormatNumber(gvr.Cells(2).Text, formatNu, TriState.True, TriState.False, TriState.True)
            ''  gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            GridView3.FooterRow.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView3.FooterRow.Cells(2).HorizontalAlign = HorizontalAlign.Center

            'gvr.Cells(4).Text = FormatNumber(gvr.Cells(4).Text, formatNu, TriState.True, TriState.False, TriState.True)
            '' gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            'GridView3.FooterRow.Cells(4).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            'GridView3.FooterRow.Cells(4).HorizontalAlign = HorizontalAlign.Center
        Next

    End Sub

    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        '
        'Add-Start by Joy 250605
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If InStr(e.Row.Cells(7).Text, "QC退件") > 0 Then
                For i = 0 To 8
                    'e.Row.Cells(i).Attributes.Add("style", "border:1px solid red ")
                    e.Row.Cells(i).BackColor = Color.Yellow
                Next
                '
                e.Row.Cells(7).ForeColor = Color.Red
                'e.Row.Cells(7).Font.Bold = True
            End If
        End If
        'Add-End

        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "合計"
        End If

    End Sub
End Class

