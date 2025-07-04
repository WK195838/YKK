Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class ComplaintOutSchedule
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
            DataList1()
            DataList2()
            DataList3()
            DataList4()

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



        sql = "   select a.step,substring(b.stepname,9,len(b.stepname)-1)stepname,sum(allcase)allcase,allcaseURL,sum(delayCase)delayCase,delaycaseURL,a.Leadtime "
        sql = sql + "  from ("
        sql = sql + "  select a.formno,a.step,substring(stepname,9,len(stepname)-1)stepname, "
        sql = sql + "  case when allcase is null then 0 else allcase end as allcase, "
        sql = sql + " case when allcase is null then '' else '...' end as allcaseURL, "
        sql = sql + "  case when delaycase  is null then  0 else delaycase end as delaycase, "
        sql = sql + "  case when delaycase is null then ''  else '...'  end as delaycaseURL,Leadtime from ("
        sql = sql + "  select a.*,hours,Hours/8/60 as Leadtime from( "
        sql = sql + "  select distinct formno,step,stepname  from M_flow a where formno = 003109 "
        sql = sql + "  and step not in (1,999,40,15,35,31,41)  )a,M_leadtime b "
        sql = sql + "  where(a.step = b.step And a.formno = b.formno) )a left join ( "
        sql = sql + "  select step,stepname as stepname1,sum(allcase)allcase,sum(delaycase)delaycase from"
        sql = sql + "  ( select step,substring(stepname,9,len(stepname)-1)stepname,1 as allcase, "
        sql = sql + "  case when delaysts ='正常' then  0 else 1 end as delaycase  from   V_WaitHandle_01"
        sql = sql + "  where formno = '003109' and flowtype ='1' and active ='1'  and step <> '40'"
        sql = sql + "  and seqno=1 )a group by step,stepname  ) b on a.step =b.step"
        sql = sql + "  )a,m_flow b"
        sql = sql + "  where(a.step = b.step And Action = 0 And a.formno = b.formno)"
        sql = sql + "  group by  a.step,a.allcaseURL,a.delaycaseURL,a.Leadtime,b.stepname"
        sql = sql + "  order by a.step"





        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
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
                h1.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If



            If e.Row.Cells(4).Text <> "0" Then
                h2.Target = "_blank"
                h2.Text = e.Row.Cells(5).Text
                ' 連結到延遲
                'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                h2.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                e.Row.Cells(5).Text = ""
                e.Row.Cells(5).Controls.Add(h2)
            End If
        
        End If
        '連結

        e.Row.Cells(6).Visible = False

    End Sub


    Sub DataList1()
        Dim sql As String

        sql = "  select * from ("
        sql = sql + " select a.step, stepname,isnull(allcase,0) allcase,"
        sql = sql + " case when allcase is null or allcase =0 then '' else '...' end as allcaseURL,"
        sql = sql + " isnull(delaycase,0)delaycase,"
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL,"
        sql = sql + " case when a.step =35 then '30' when a.step =40 then '10' else '' end as leadtime"
        sql = sql + " from ("
        sql = sql + " select step,stepname as stepname,sum(allcase)allcase,sum(delaycase)delaycase from ( "
        sql = sql + " select step,substring(stepname,9,len(stepname)-1)stepname,"
        sql = sql + " case when step in (35,40) then 1 "
        sql = sql + " when step =999 and sts =1 then 1 else 0 end as allcase, case when delaysts ='正常' then  0 else 1 end as delaycase"
        sql = sql + " from   V_WaitHandle_01  where formno = '003109' and flowtype ='1' and active ='1'  and step in('40','35') and seqno=1 "
        sql = sql + " )a group by step,stepname "
        sql = sql + " )a "
        sql = sql + " )a order by step"
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView2.DataSource = dtData
        GridView2.DataBind()
    End Sub

    Sub DataList2()
        Dim sql As String

        sql = "  select * from ("
        sql = sql + " select a.step, stepname,isnull(allcase,0) allcase,"
        sql = sql + " case when allcase is null or allcase =0 then '' else '...' end as allcaseURL,"
        sql = sql + " isnull(delaycase,0)delaycase,"
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL,"
        sql = sql + " case when a.step =35 then '30' when a.step =40 then '10' else '' end as leadtime"
        sql = sql + " from ("
        sql = sql + " select step,stepname as stepname,sum(allcase)allcase,sum(delaycase)delaycase from ( "
        sql = sql + " select step,'完了' stepname,"
        sql = sql + " case when step in (35,40) then 1 "
        sql = sql + " when a.step =999 and b.sts =1 then 1 else 0 end as allcase, case when delaysts ='正常' then  0 else 1 end as delaycase"
        sql = sql + " from    V_WaitHandle_01 a,F_ComplaintOutSheet b"
        sql = sql + " where  a.formsno = b.formsno "
        sql = sql + " and a.formno = '003109' and step ='999'"
        sql = sql + " and b.sts =1"
        sql = sql + " )a group by step,stepname "
        sql = sql + " )a "
        sql = sql + " )a order by step"
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView3.DataSource = dtData
        GridView3.DataBind()
    End Sub

    Sub DataList3()
        Dim sql As String



        sql = "   select * from ("
        sql = sql + "  select a.step, stepname,isnull(allcase,0) allcase, case when allcase is null or allcase =0 then '' else '...' end as allcaseURL, isnull(delaycase,0)delaycase, "
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL, case when a.step =35 then '30' when a.step =40 then '10' else '' end as leadtime from ( "
        sql = sql + " select step,stepname as stepname,sum(allcase)allcase,sum(delaycase)delaycase from (  "
        sql = sql + " select a.step,a.stepname,isnull(allcase,0)allcase,isnull(delaycase,0)delaycase  from ("
        sql = sql + " select distinct step,'完了' stepname  from  "
        sql = sql + " V_WaitHandle_01 a,F_ComplaintOutSheet b where  a.formsno = b.formsno  and a.formno = '003109'    and step ='999'  "
        sql = sql + " and b.sts =1"
        sql = sql + " )a left join ("
        sql = sql + " select step,'完了' stepname, case when step in (35,40) then 1  "
        sql = sql + " when a.step =999 and b.sts =1 then 1 else 0 end as allcase, case when delaysts ='正常' then  0 else 1 end as delaycase from  "
        sql = sql + " V_WaitHandle_01 a,F_ComplaintOutSheet b where  a.formsno = b.formsno  and a.formno = '003109' and  rely =1  and step ='999'  "
        sql = sql + " and b.sts =1"
        sql = sql + " )b on a.step =b.step"
        sql = sql + " "
        sql = sql + " )a group by step,stepname  )a "
        sql = sql + "  )a order by step"


        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView4.DataSource = dtData
        GridView4.DataBind()
    End Sub

    Sub DataList4()
        Dim sql As String



        sql = "   select * from ("
        sql = sql + "  select a.step, stepname,isnull(allcase,0) allcase, case when allcase is null or allcase =0 then '' else '...' end as allcaseURL, isnull(delaycase,0)delaycase, "
        sql = sql + " case when delaycase is null or delaycase =0 then ''  else '...'  end as delaycaseURL, case when a.step =35 then '30' when a.step =40 then '10' else '' end as leadtime from ( "
        sql = sql + " select step,stepname as stepname,sum(allcase)allcase,sum(delaycase)delaycase from (  "
        sql = sql + " select a.step,a.stepname,isnull(allcase,0)allcase,isnull(delaycase,0)delaycase  from ("
        sql = sql + " select distinct step,'完了' stepname  from  "
        sql = sql + " V_WaitHandle_01 a,F_ComplaintOutSheet b where  a.formsno = b.formsno  and a.formno = '003109'    and step ='999'  "
        sql = sql + " and b.sts =1"
        sql = sql + " )a left join ("
        sql = sql + " select step,'完了' stepname, case when step in (35,40) then 1  "
        sql = sql + " when a.step =999 and b.sts =1 then 1 else 0 end as allcase, case when delaysts ='正常' then  0 else 1 end as delaycase from  "
        sql = sql + " V_WaitHandle_01 a,F_ComplaintOutSheet b where  a.formsno = b.formsno  and a.formno = '003109' and  mat =1  and step ='999'  "
        sql = sql + " and b.sts =1"
        sql = sql + " )b on a.step =b.step"
        sql = sql + " "
        sql = sql + " )a group by step,stepname  )a "
        sql = sql + "  )a order by step"



        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView5.DataSource = dtData
        GridView5.DataBind()
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

    Protected Sub GridView2_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.DataBound
        Dim formatNu As Integer

        formatNu = 0

        Dim intTotal As Decimal = 0
        Dim intTotal1 As Decimal = 0
        For Each gvr As GridViewRow In GridView2.Rows


            intTotal += gvr.Cells(2).Text

            intTotal1 += gvr.Cells(4).Text

            gvr.Cells(2).Text = FormatNumber(gvr.Cells(2).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(2).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(3).HorizontalAlign = HorizontalAlign.Center
            GridView2.FooterRow.Cells(2).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)
            GridView2.FooterRow.Cells(2).HorizontalAlign = HorizontalAlign.Center

            gvr.Cells(4).Text = FormatNumber(gvr.Cells(4).Text, formatNu, TriState.True, TriState.False, TriState.True)
            'gvr.Cells(4).HorizontalAlign = HorizontalAlign.Center
            'gvr.Cells(5).HorizontalAlign = HorizontalAlign.Center
            GridView2.FooterRow.Cells(4).Text = FormatNumber(intTotal1, formatNu, TriState.True, TriState.False, TriState.True)
            GridView2.FooterRow.Cells(4).HorizontalAlign = HorizontalAlign.Center
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
                h1.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If



            If e.Row.Cells(4).Text <> "0" Then
                h2.Target = "_blank"
                h2.Text = e.Row.Cells(5).Text
                ' 連結到延遲
                'h2.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                h2.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pDelay=1"
                e.Row.Cells(5).Text = ""
                e.Row.Cells(5).Controls.Add(h2)
            End If
        End If
        '連結

        e.Row.Cells(6).Visible = False
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
        LFun01.NavigateUrl = "ComplaintOutSheet_01.aspx?pFormNo=003109" & _
                                                          "&pFormSno=0" & _
                                                          "&pStep=" & xStep & _
                                                          "&pSeqNo=0" & _
                                                          "&pUserID=" & Request.QueryString("pUserID") & _
                                                          "&pApplyID=" & Request.QueryString("pUserID")
        LFun01.Target = "_blank"

        LFun02.Enabled = True
        LFun02.NavigateUrl = "ComplaintOutinqCommission.aspx"

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

    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound


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
                h1.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pRemark=0"
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If

          
        End If
        '連結

        e.Row.Cells(6).Visible = False
    End Sub

    Protected Sub GridView4_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowDataBound

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
                h1.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pRemark=1"
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If


        End If
        '連結

        e.Row.Cells(6).Visible = False
    End Sub

    Protected Sub GridView5_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView5.RowDataBound

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
                h1.NavigateUrl = "ComplaintOutScheduleAll.aspx?&pStep=" + wStep + "&pRemark=2"
                e.Row.Cells(3).Text = ""
                e.Row.Cells(3).Controls.Add(h1)
            End If


        End If
        '連結

        e.Row.Cells(6).Visible = False
    End Sub
End Class

