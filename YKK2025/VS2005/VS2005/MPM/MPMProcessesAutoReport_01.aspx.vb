Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class MPMProcessesAutoReport_01
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String = "CL"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

    Dim wUserName As String = ""    '姓名代理用
    Dim HolidayList As New List(Of Integer) '用以記錄假日的欄位索引值
    Dim indexList As New List(Of Integer)   '用以記錄不屬於選取月份的欄位索引值
    Dim DateList As New List(Of String)     '

    ''' <summary>
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer
    Dim Nodata As Integer
    Dim cun, cun1 As Integer
    Dim c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 As Integer
    Dim NewPage, TotPage As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()          '設定共用參數
        GridView1.Style("top") = 120 & "px"

        If Not Me.IsPostBack Then   '不是PostBack
            ' Label7.Style.Add("text-align", "right")
            ' Label12.Style.Add("text-align", "left")
            Label7.Text = "1"

            GetEngine()
            ShowFormData()      '顯示表單資料
            Marquee() '跑馬燈
        End If

    End Sub

    Sub Marquee()


        'obj.Text = "<MARQUEE onmouseover=stop(); onmouseout=start();  text-align:left scrollAmount=3 scrollDeay=100 direction=right  height=340 width=200 scrolltop=0 scrollleft=0>  "
        'obj.Text += "<table width='100%' bordr='1' cellspacing='6' cellpaddig='6'>"
        '    obj.Text = " <marquee direction=up scrollamount=1 scrolldelay=100 onmouseover='this.stop()' onmouseout='this.start()' direction=up  height=400 width=230 >"
        obj.Text = " <marquee bgcolor='#000000' border='0' align='left' scrollamount='9'  direction='left'    width='980' height='48' > "

        Dim sql As String
        Dim EngineStr As String = ""
        sql = "select   * from F_Marquee where action =1  "
        sql = sql + " and  ( startdate   ='1900/01/01' or "
        sql = sql + " convert(char(10),startdate,111) <= convert(char(10),getdate(),111) )"
        sql = sql + " and  ( Enddate   ='1900/01/01' or  convert(char(10),enddate,111) >= convert(char(10),getdate(),111) )"
        sql = sql + " order by Unique_id "


        Dim dt1 As Data.DataTable = uDataBase.GetDataTable(sql)
        For Each dtr1 As Data.DataRow In dt1.Rows

            obj.Text += "<tr><td  align='left'>"
            '     obj.Text += "<img src='/MinsuHome/images/a00_new-top06_d.gif'  border='0'>"
            '      obj.Text += "</td><td style='FONT-SIZE: 11px'>"
            '  obj.Text += "<a href='/MinsuHome/Activity/ViewActivity.aspx?ID=" & dtr1("formsno").ToString & "'>"
            obj.Text += " <font Size=" + dtr1("FontSize") + "' color='" + dtr1("color") + "' face='" + dtr1("FontStyle") + "'    >" & dtr1("Subject") & "</font>"
            obj.Text += "</a>"
            obj.Text += "</td></tr>"

        Next

        obj.Text += "</table>"
        obj.Text += "</marquee>"
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID

        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        Response.Cookies("PGM").Value = "MPMProcessesReport_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

    End Sub


    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub GetEngine()
        DFinish.SelectedIndex = 1

        Dim sql As String
        Dim i As Integer
        sql = " select  unique_id, rtrim(Data)Data  from m_referp"
        sql = sql + " where  cat = '4001'"
        sql = sql + " and dkey = 'EngineSelect-單獨選取' order by unique_id"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)
        DEngine.Items.Clear()
        DEngine.Items.Add("全部")
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = Trim(DBAdapter1.Rows(i).Item("Data"))
            ListItem1.Value = Trim(DBAdapter1.Rows(i).Item("Data"))
            DEngine.Items.Add(ListItem1)
        Next
    End Sub





    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim EngineSQL As String = ""
        Dim EngineName As String = ""
        Dim StsSQL As String = ""
        Dim SdelaySQL As String = ""
        Dim Sts As String = ""
        Dim type1 As String = ""
        Dim type2 As String = ""
        Dim Finish As String = ""
        Dim mapno As String = ""

        If CompareValidator1.IsValid = False Then '設定秒數
            Timer1.Interval = "60000"
            DSecond.Text = "60"
        Else
            Timer1.Interval = Int(DSecond.Text) * 1000
        End If



        If DType1.SelectedValue <> "全部" Then
            type1 = " and Type1='" + Trim(DType1.SelectedValue) + "'"
        End If

        If DType2.SelectedValue <> "全部" Then
            type2 = " and Type2='" + Trim(DType2.SelectedValue) + "'"
        End If

        If DFinish.SelectedItem.Text <> "全部" Then

            Finish = " and Sts= " + DFinish.SelectedValue


        End If


        If DMapNo.Text <> "" Then
            mapno = " and MapNo like '%" + DMapNo.Text + "%'"
        End If


        EngineName = DEngine.SelectedValue
        Sts = DSts.SelectedValue

        If DEngine.SelectedValue <> "全部" And Sts <> "全部" Then
            EngineSQL = " and  ( ( substring(engine1,1,1) = '" + Sts + "' and   substring(engine1,3,len(engine1)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine2,1,1) = '" + Sts + "' and   substring(engine2,3,len(engine2)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine3,1,1) = '" + Sts + "' and   substring(engine3,3,len(engine3)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine4,1,1)= '" + Sts + "' and   substring(engine4,3,len(engine4)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " ( substring(engine5,1,1) = '" + Sts + "' and   substring(engine5,3,len(engine5)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine6,1,1)= '" + Sts + "' and  substring(engine6,3,len(engine6)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " ( substring(engine7,1,1)= '" + Sts + "' and substring(engine7,3,len(engine7)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine8,1,1) = '" + Sts + "' and substring(engine8,3,len(engine8)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " ( substring(engine9,1,1) = '" + Sts + "' and  substring(engine9,3,len(engine8)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine10,1,1) = '" + Sts + "' and  substring(engine10,3,len(engine10)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " ( substring(engine11,1,1) = '" + Sts + "' and  substring(engine11,3,len(engine11)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine12,1,1) = '" + Sts + "' and  substring(engine12,3,len(engine12)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " ( substring(engine13,1,1) = '" + Sts + "' and  substring(engine13,3,len(engine13)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine14,1,1) = '" + Sts + "' and   substring(engine14,3,len(engine14)-1) =  '" + EngineName + "' ) ) "
        ElseIf DEngine.SelectedValue <> "全部" And Sts = "全部" Then
            EngineSQL = " and  ( (   substring(engine1,3,nullif(len(engine1),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (   substring(engine2,3,nullif(len(engine2),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (   substring(engine3,3,nullif(len(engine3),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (   substring(engine4,3,nullif(len(engine4),0)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " (    substring(engine5,3,nullif(len(engine5),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (  substring(engine6,3,nullif(len(engine6),0)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " (  substring(engine7,3,nullif(len(engine7),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " ( substring(engine8,3,nullif(len(engine8),0)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " (  substring(engine9,3,nullif(len(engine9),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (  substring(engine10,3,nullif(len(engine10),0)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " (   substring(engine11,3,nullif(len(engine11),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (  substring(engine12,3,nullif(len(engine12),0)-1) =  '" + EngineName + "' ) Or "
            EngineSQL = EngineSQL + " (   substring(engine13,3,nullif(len(engine13),0)-1) =  '" + EngineName + "' ) or"
            EngineSQL = EngineSQL + " (    substring(engine14,3,nullif(len(engine14),0)-1) =  '" + EngineName + "' ) ) "
        End If


        If DSts.SelectedValue <> "全部" Then
            StsSQL = " and ( left(engine1,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine2,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine3,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine4,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine5,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine6,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine7,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine8,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine9,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine10,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine11,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine12,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine13,1) = '" + Sts + "' or "
            StsSQL = StsSQL + "left(engine14,1) = '" + Sts + "' ) "

        Else
            StsSQL = " and ( engine1 = '' or "
            StsSQL = StsSQL + " engine2 = '' or "
            StsSQL = StsSQL + " engine3 = '' or "
            StsSQL = StsSQL + " engine4 = '' or "
            StsSQL = StsSQL + " engine5 = '' or "
            StsSQL = StsSQL + " engine6 = '' or "
            StsSQL = StsSQL + " engine7 = '' or "
            StsSQL = StsSQL + " engine8 = '' or "
            StsSQL = StsSQL + " engine9 = '' or "
            StsSQL = StsSQL + " engine10 = '' or "
            StsSQL = StsSQL + " engine11 = '' or "
            StsSQL = StsSQL + " engine12 = '' or "
            StsSQL = StsSQL + " engine13 = '' or "
            StsSQL = StsSQL + " engine14 = '' ) "


        End If



        If DSdelay.SelectedValue = "正常" Then
            SdelaySQL = " and  Sdelay  =0 "
        ElseIf DSdelay.SelectedValue = "遲納" Then
            SdelaySQL = " and  Sdelay  =1 "
        Else
            SdelaySQL = ""
        End If


        Dim SQL As String
        SQL = " select *,case when sdelay = 1 then  datediff(day,getdate(),finishdate) else 0 end as delayday from ("
        SQL = SQL + " select *  from V_MPMProcessesSheet  where 1=1  "
        SQL = SQL + EngineSQL + SdelaySQL + type1 + type2 + Finish + mapno




        ' 加上排序 
        SQL = SQL + ")a "
        SQL = SQL + " where 1=1 " + StsSQL
        SQL = SQL + " ORDER BY " + DField.SelectedValue
        If DOrderby.SelectedValue = "降順" Then
            SQL = SQL + " desc "
        End If

        Nodata = 1

        GridView1.PageSize = Int(DCount.Text)

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        TotPage = DBAdapter1.Rows.Count


        If DBAdapter1.Rows.Count > 0 Then

            GridView1.DataSource = DBAdapter1
            GridView1.DataBind()

            Nodata = 1
        Else

            Nodata = 0
            GridView1.DataSource = Nothing
            GridView1.DataBind()

        End If
        Label13.Text = TotPage
        Label12.Text = "共" + Str(GridView1.PageCount) + "頁"


    End Sub

    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Response.AppendHeader("Content-Disposition", "attachment;filename=MPM_ProcessesRemport.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
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



    Protected Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        ' Button2_Click(sender, e) 觸發button
        If Timer1.Enabled = True Then
            If Me.GridView1.PageCount > 1 Then
                If Me.GridView1.PageIndex = Me.GridView1.PageCount - 1 Then
                    Me.GridView1.PageIndex = 0
                    ShowFormData()

                Else




                    If Me.GridView1.PageCount = Int(Label7.Text) Then  '如果到最後一頁就重頭開始
                        Me.GridView1.PageIndex = 0
                    Else
                        Me.GridView1.PageIndex = Me.GridView1.PageIndex + 1
                    End If

                    '如果有變換條件就重頭開始
                    If c1 <> 0 Or c2 <> 0 Or c3 <> 0 Or c4 <> 0 Or c5 <> 0 Or c6 <> 0 Or c7 <> 0 Or c8 <> 0 Or c9 <> 0 Or c10 <> 0 Then
                        Me.GridView1.PageIndex = 0
                        Label7.Text = "1"
                    Else
                        Label7.Text = Me.GridView1.PageIndex + 1
                        Label7.Text = Trim(Label7.Text)
                        NewPage = Int(Label7.Text)
                        cun1 = Label7.Text
                    End If
                    ShowFormData()


                    If Me.GridView1.PageCount = NewPage Then
                        Me.GridView1.PageIndex = 0
                        NewPage = 1
                    End If




                    c1 = 0
                    c2 = 0
                    c3 = 0
                    c4 = 0
                    c5 = 0
                    c6 = 0
                    c7 = 0
                    c8 = 0
                    c9 = 0
                    c10 = 0
                End If
            End If
        End If



    End Sub



    Protected Sub GridView1_RowCreated1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        If (e.Row.RowType = DataControlRowType.Header) Then
            ' 建立自訂的標題 

            Dim gv As GridView = DirectCast(sender, GridView)
            Dim gvRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            ' 增加欄位 

            Dim tc1 As New TableCell()
            tc1.Text = "加工編號"
            gvRow.Cells.Add(tc1)


            Dim tc2 As New TableCell()
            tc2.Text = "圖號"
            gvRow.Cells.Add(tc2)

            Dim tc3 As New TableCell()
            tc3.Text = "收件日"
            gvRow.Cells.Add(tc3)

            Dim tc4 As New TableCell()
            tc4.Text = "預定完成日"
            'tc4.ColumnSpan = 2
            gvRow.Cells.Add(tc4)


            Dim tc5 As New TableCell()
            tc5.Text = "遲納<br/>天數"
            'tc4.ColumnSpan = 2
            gvRow.Cells.Add(tc5)

            Dim tc6 As New TableCell()
            tc6.Text = "部門"
            gvRow.Cells.Add(tc6)


            Dim tc7 As New TableCell()
            tc7.Text = "類別"
            tc7.ColumnSpan = 2   ' 跨二欄
            gvRow.Cells.Add(tc7)




            Dim tc9 As New TableCell()
            tc9.Text = "工程"
            tc9.ColumnSpan = 14   ' 跨二欄
            gvRow.Cells.Add(tc9)



            e.Row.Cells.Clear()
            gv.Controls(0).Controls.AddAt(0, gvRow)

        End If
    End Sub

    Protected Sub GridView1_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        '  If e.Row.Cells(2).Text = "合計" Then
        'e.Row.Cells(2).Text = "合計"
        If e.Row.RowType <> DataControlRowType.Header Then

            e.Row.Cells(0).Attributes.Add("style", "vnd.ms-excel.numberformat:@") '轉成文字格式
            Dim i As Integer
            Dim h1 As New HyperLink
            h1.Text = e.Row.Cells(0).Text
            h1.NavigateUrl = "MPMProcessesReport_02.aspx?&pNo=" & e.Row.Cells(0).Text
            h1.Target = "_blank"
            e.Row.Cells(0).Text = ""
            e.Row.Cells(0).Controls.Add(h1)

            Dim drv As DataRowView = CType(e.Row.DataItem, DataRowView)


            If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then


                If Not drv Is Nothing Then

                    If e.Row.Cells(3).Text < Now.ToString("yyyy/MM/dd") Then
                        For i = 0 To 4
                            e.Row.Cells(i).ForeColor = Color.Red
                            e.Row.Cells(i).Font.Bold = True

                        Next
                    End If

                    For i = 8 To 19 '判斷顏色

                        e.Row.Cells(i).HorizontalAlign = HorizontalAlign.Center


                        If Mid(e.Row.Cells(i).Text, 1, 1) = "R" Then
                            e.Row.Cells(i).BackColor = Drawing.Color.LightBlue
                        ElseIf Mid(e.Row.Cells(i).Text, 1, 1) = "S" Then
                            e.Row.Cells(i).BackColor = Drawing.Color.Yellow
                        ElseIf Mid(e.Row.Cells(i).Text, 1, 1) = "E" Then
                            e.Row.Cells(i).BackColor = Drawing.Color.LightGreen
                        ElseIf Mid(e.Row.Cells(i).Text, 1, 1) = "X" Then
                            e.Row.Cells(i).BackColor = Drawing.Color.Red
                        ElseIf Mid(e.Row.Cells(i).Text, 1, 1) = "A" Then
                            e.Row.Cells(i).BackColor = Drawing.Color.LightGray
                        End If
                        If e.Row.Cells(i).Text <> "&nbsp;" Then

                            e.Row.Cells(i).Text = Mid(e.Row.Cells(i).Text, 3, Len(e.Row.Cells(i).Text) - 1)
                        End If

                    Next

                End If

            End If


        End If



    End Sub

    Protected Sub GridView1_PageIndexChanging1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        GridView1.PageIndex = e.NewPageIndex
        ShowFormData()
    End Sub



    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click

        If Timer1.Enabled = True Then



        Else


        End If
        If Button4.Text = "進行" Then
            Timer1.Enabled = True
            Button4.Text = "暫停"
            Button4.BackColor = Color.Chartreuse

            Button2.Enabled = False
            Button3.Enabled = False
        Else

            Timer1.Enabled = False
            Button4.Text = "進行"
            Button4.BackColor = Color.PeachPuff
            Button2.Enabled = True
            Button3.Enabled = True
        End If
        Me.GridView1.PageIndex = CInt(Label7.Text) - 1
        ShowFormData()
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If c1 <> 0 Or c2 <> 0 Or c3 <> 0 Or c4 <> 0 Or c5 <> 0 Or c6 <> 0 Or c7 <> 0 Or c8 <> 0 Or c9 <> 0 Or c10 <> 0 Then
            Me.GridView1.PageIndex = 0
            Label7.Text = "1"
        Else
            Label7.Text = Me.GridView1.PageIndex + 1
            Label7.Text = Trim(Label7.Text)
            NewPage = Int(Label7.Text)
            cun1 = Label7.Text
        End If

        ShowFormData()
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub Button2_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click

        If CInt(Label7.Text) - 1 > 0 Then
            Me.GridView1.PageIndex = CInt(Label7.Text) - 2
            Label7.Text = CInt(Label7.Text) - 1

            If Label7.Text = 1 Then
                Button2.Enabled = False
            End If
            ShowFormData()
        End If

        '  ShowFormData()
    End Sub

    Protected Sub Button3_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button2.Enabled = True
        Me.GridView1.PageIndex = CInt(Label7.Text) + 1
        Label7.Text = CInt(Label7.Text) + 1
        ShowFormData()
        ' ShowFormData()
    End Sub

    Protected Sub DEngine_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DEngine.TextChanged
        c1 = 1
    End Sub

    Protected Sub DSts_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSdelay.TextChanged
        c2 = 1
    End Sub

    Protected Sub DFinish_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFinish.TextChanged
        c3 = 1
    End Sub

    Protected Sub DType1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DType1.TextChanged
        c4 = 1
    End Sub

    Protected Sub DType2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DType2.TextChanged
        c5 = 1
    End Sub

    Protected Sub DField_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DField.TextChanged
        c5 = 1
    End Sub

    Protected Sub DOrderby_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DOrderby.TextChanged
        c7 = 1
    End Sub

    Protected Sub DMapNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMapNo.TextChanged
        c8 = 1
    End Sub

    Protected Sub DSecond_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSecond.TextChanged
        c9 = 1
    End Sub

    Protected Sub DCount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCount.TextChanged
        c10 = 1
    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BStart.Click
       
        If BStart.Text = "展開" Then
            GridView1.Style("top") = 290 & "px"
            PManu.Visible = True
            PManu.Style("top") = 140 & "px"
            BStart.Text = "收回"
            BStart.BackColor = Color.Violet
            GridView1.DataBind()

        Else
            BStart.Text = "展開"
            GridView1.Style("top") = 120 & "px"
            PManu.Visible = False
            BStart.BackColor = Color.Aqua
        End If
        '   ShowFormData()
        
    End Sub

  
End Class

