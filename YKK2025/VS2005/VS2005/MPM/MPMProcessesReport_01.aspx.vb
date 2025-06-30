Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb



Partial Class MPMProcessesReport_01
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '�U���
    Dim Attribute(60) As Integer    '�U����ݩ�    
    Dim Top As String               '�w�]�������m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim wAgentID As String          '�Q�N�z�HID
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim wNextGate As String         '�U�@���H
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '�x�_��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End

    Dim wUserName As String = ""    '�m�W�N�z��
    Dim HolidayList As New List(Of Integer) '�ΥH�O�����骺�����ޭ�
    Dim indexList As New List(Of Integer)   '�ΥH�O�����ݩ�������������ޭ�
    Dim DateList As New List(Of String)     '

    ''' <summary>
    ''' �H�U���@�Ψ禡���ŧi
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer
    Dim Nodata As Integer
    Dim cun, cun1 As Integer


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()          '�]�w�@�ΰѼ�


        If Not Me.IsPostBack Then   '���OPostBack

            GetEngine()
            ShowFormData()      '��ܪ����

        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID

        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        Response.Cookies("PGM").Value = "MPMProcessesReport_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X

        DSdate.Text = Format(Now.Date, "yyyy/MM/dd")

    End Sub


    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub GetEngine()
        DFinish.SelectedIndex = 1

        Dim sql As String
        Dim i As Integer
        sql = " select  unique_id, rtrim(Data)Data  from m_referp"
        sql = sql + " where  cat = '4001'"
        sql = sql + " and dkey = 'EngineSelect-��W���' order by unique_id"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)
        DEngine.Items.Clear()
        DEngine.Items.Add("����")
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = Trim(DBAdapter1.Rows(i).Item("Data"))
            ListItem1.Value = Trim(DBAdapter1.Rows(i).Item("Data"))
            DEngine.Items.Add(ListItem1)
        Next
    End Sub





    '*****************************************************************
    '**
    '**     ��ܪ����
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
        Dim clinterSQL As String = ""
        Dim SQL1 As String = ""
        Dim Sdate As String = ""
        Dim Edate As String = ""
        Dim DateSQL As String = ""
        Dim DelaySQL As String = ""

        If DType1.SelectedValue <> "����" Then
            type1 = " and Type1='" + Trim(DType1.SelectedValue) + "'"
        End If

        If DType2.SelectedValue <> "����" Then
            type2 = " and Type2='" + Trim(DType2.SelectedValue) + "'"
        End If

        If DFinish.SelectedItem.Text <> "����" Then

            Finish = " and Sts= " + DFinish.SelectedValue
        End If


        If DMapNo.Text <> "" Then
            mapno = " and MapNo like '%" + DMapNo.Text + "%'"
        End If

        '�ӽФ�
        If IsDate(BSdate.Text) = True And IsDate(DSdate.Text) = True Then


            Sdate = BSdate.Text
            Edate = DSdate.Text
            DateSQL = DateSQL + " And AppDate between '" + Sdate + "'  and '" + Edate + "'"
        End If




        EngineName = DEngine.SelectedValue
        Sts = DSts.SelectedValue

        If DEngine.SelectedValue <> "����" And Sts <> "����" Then
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
        ElseIf DEngine.SelectedValue <> "����" And Sts = "����" Then
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
        'ElseIf DEngine.SelectedValue <> "����" And DSdelay.SelectedValue = "0" Then
        ' EngineSQL = " and  ( ( engine1 <>'' and   substring(engine1,3,len(engine1)-1) =  '" + EngineName + "' ) or"
        ' EngineSQL = EngineSQL + " (  engine2 <>'' and   substring(engine2,3,len(engine2)-1) =  '" + EngineName + "' ) or"
        ' EngineSQL = EngineSQL + " (  engine3 <>'' and   substring(engine3,3,len(engine3)-1) =  '" + EngineName + "' ) or"
        ' EngineSQL = EngineSQL + " (  engine4 <>''  and   substring(engine4,3,len(engine4)-1) =  '" + EngineName + "' ) Or "
        ' EngineSQL = EngineSQL + " (  engine5 <>'' and   substring(engine5,3,len(engine5)-1) =  '" + EngineName + "' ) or"
        ' EngineSQL = EngineSQL + " (  engine6 <>'' and  substring(engine6,3,len(engine6)-1) =  '" + EngineName + "' ) Or "
        ' EngineSQL = EngineSQL + " (  engine7 <>'' and substring(engine7,3,len(engine7)-1) =  '" + EngineName + "' ) or"
        ' EngineSQL = EngineSQL + " (  engine8 <>'' and substring(engine8,3,len(engine8)-1) =  '" + EngineName + "' ) Or "
        ' EngineSQL = EngineSQL + " (  engine9 <>'' and  substring(engine9,3,len(engine9)-1) =  '" + EngineName + "' ) or"
        ' EngineSQL = EngineSQL + " (  engine10 <>'' and  substring(engine10,3,len(engine10)-1) =  '" + EngineName + "' ) Or "
        ' EngineSQL = EngineSQL + " (  engine11 <>'' and  substring(engine11,3,len(engine11)-1) =  '" + EngineName + "' ) or"
        '  EngineSQL = EngineSQL + " (  engine12 <>'' and  substring(engine12,3,len(engine12)-1) =  '" + EngineName + "' ) Or "
        '   EngineSQL = EngineSQL + " (  engine13 <>'' and  substring(engine13,3,len(engine13)-1) =  '" + EngineName + "' ) or"
        '    EngineSQL = EngineSQL + " (  engine14 <>'' and   substring(engine14,3,len(engine14)-1) =  '" + EngineName + "' ) ) "
        ' End If

        If DSts.SelectedValue <> "����" Then
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
            StsSQL = " "


        End If

        If DDelay.SelectedValue <> "����" Then
            DelaySQL = " and ( delay1 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay2 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay3 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay4 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay5 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay6 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay7 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay8 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay9 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay10 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay11 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay12 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay13 ='" + DDelay.SelectedValue + "' or "
            DelaySQL = DelaySQL + "delay14 ='" + DDelay.SelectedValue + "' ) "
        End If

        If DSdelay.SelectedValue = "���`" Then
            SdelaySQL = " and  Sdelay  =0 "
        ElseIf DSdelay.SelectedValue = "���" Then
            SdelaySQL = " and  Sdelay  =1 "
        Else
            SdelaySQL = ""
        End If

        If DClinter.Text <> "" Then
            clinterSQL = " and clinter  like '%" + DClinter.Text + "%'"
        Else
            clinterSQL = ""
        End If


        Dim SQL As String
        SQL = " select  *,case when sts =1 then datediff(day,afinishdate,finishdate) else datediff(day,getdate(),finishdate) end as delayday from ("
        SQL = SQL + " select *  from V_MPMProcessesSheet  where 1=1  "
        SQL = SQL + EngineSQL + SdelaySQL + type1 + type2 + Finish + mapno + clinterSQL + DateSQL + DelaySQL


        ' �[�W�Ƨ� 
        SQL = SQL + ")a "
        SQL = SQL + " where 1=1 " + StsSQL

        '�p�ⵧ��
        If DFinish.SelectedItem.Text = "����" Then
            SQL1 = " select count(*)Cun from (" + SQL + ")a"
            SQL = " select  * from (" + SQL + ")a"
        End If


        SQL = SQL + " ORDER BY " + DField.SelectedValue
        If DOrderby.SelectedValue = "����" Then
            SQL = SQL + " desc "
        End If

        Nodata = 1

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            LAutoReport.NavigateUrl = "MPMProcessesAutoReport_01.aspx?&pEngine=" + DBAdapter1.Rows(0).Item("No")

            GridView1.DataSource = DBAdapter1
            GridView1.DataBind()

            Nodata = 1
        Else

            Nodata = 0
            GridView1.DataSource = Nothing
            GridView1.DataBind()

        End If

        If DFinish.SelectedItem.Text = "����" Then
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL1)
            Label7.Text = "���ơG" + Str(DBAdapter2.Rows(0).Item("Cun"))
        Else
            Label7.Text = "���ơG" + Str(DBAdapter1.Rows.Count)
        End If






    End Sub

    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Response.AppendHeader("Content-Disposition", "attachment;filename=MPM_ProcessesRemport.xls")     '�{���O���P
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
    '**     ��Excel�@�ε{��
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub



    Protected Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        ' Button2_Click(sender, e) Ĳ�obutton
        ShowFormData()
    End Sub



    Protected Sub GridView1_RowCreated1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        If (e.Row.RowType = DataControlRowType.Header) Then
            ' �إߦۭq�����D 

            Dim gv As GridView = DirectCast(sender, GridView)
            Dim gvRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

            ' �W�[��� 

            Dim tc1 As New TableCell()
            tc1.Text = "�[�u�s��"
            gvRow.Cells.Add(tc1)


            Dim tc2 As New TableCell()
            tc2.Text = "�ϸ�"
            gvRow.Cells.Add(tc2)

            Dim tc3 As New TableCell()
            tc3.Text = "�����"
            gvRow.Cells.Add(tc3)

            Dim tc4 As New TableCell()
            tc4.Text = "�w�w������"
            'tc4.ColumnSpan = 2
            gvRow.Cells.Add(tc4)


            Dim tc5 As New TableCell()
            tc5.Text = "���<Br/>�Ѽ�"
            'tc4.ColumnSpan = 2
            gvRow.Cells.Add(tc5)

            Dim tc6 As New TableCell()
            tc6.Text = "�̿��"
            gvRow.Cells.Add(tc6)


            Dim tc7 As New TableCell()
            tc7.Text = "����"
            'tc7.ColumnSpan = 2   ' ��G��
            gvRow.Cells.Add(tc7)


            Dim tc8 As New TableCell()
            tc8.Text = "���O"
            tc8.ColumnSpan = 2   ' ��G��
            gvRow.Cells.Add(tc8)



            Dim tc9 As New TableCell()
            tc9.Text = "�u�{"
            tc9.ColumnSpan = 14  ' ��G��
            gvRow.Cells.Add(tc9)


            e.Row.Cells.Clear()
            gv.Controls(0).Controls.AddAt(0, gvRow)

        End If
    End Sub

    Protected Sub GridView1_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        '  If e.Row.Cells(2).Text = "�X�p" Then
        'e.Row.Cells(2).Text = "�X�p"
        If e.Row.RowType <> DataControlRowType.Header Then

            e.Row.Cells(0).Attributes.Add("style", "vnd.ms-excel.numberformat:@") '�ন��r�榡
            Dim i, j As Integer
            Dim h1 As New HyperLink
            h1.Text = e.Row.Cells(0).Text
            h1.NavigateUrl = "MPMProcessesReport_02.aspx?&pNo=" & e.Row.Cells(0).Text
            h1.Target = "_blank"
            e.Row.Cells(0).Text = ""
            e.Row.Cells(0).Controls.Add(h1)

            Dim drv As DataRowView = CType(e.Row.DataItem, DataRowView)


            If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then


                If Not drv Is Nothing Then

                    If e.Row.Cells(4).Text < 0 Then
                        For i = 0 To 4
                            e.Row.Cells(i).ForeColor = Color.Red
                            e.Row.Cells(i).Font.Bold = True
                        Next
                    End If

                    For i = 9 To 22 '�P�_�C��

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
                            Dim a As String
                            a = e.Row.Cells(i).Text

                            If Trim(e.Row.Cells(i).Text) <> "" Then
                                If e.Row.Cells(i + 14).Text <> "&nbsp;" And e.Row.Cells(i + 14).Text <> "" And Mid(Trim(e.Row.Cells(i).Text), 1, 1) = "E" Then
                                    e.Row.Cells(i).Text = Mid(e.Row.Cells(i).Text, 3, Len(e.Row.Cells(i).Text) - 1) + "<Br/> �u�ɡG" + e.Row.Cells(i + 14).Text + "<Br/>" + e.Row.Cells(i + 28).Text

                                Else
                                    e.Row.Cells(i).Text = Mid(e.Row.Cells(i).Text, 3, Len(e.Row.Cells(i).Text) - 1)
                                End If

                            End If




                            If e.Row.Cells(i + 28).Text = "����" Then
                                e.Row.Cells(i).ForeColor = Color.Red
                            Else
                                e.Row.Cells(i).ForeColor = Color.Black
                            End If

                            e.Row.Cells(i + 14).Visible = False
                            e.Row.Cells(i + 28).Visible = False


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

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        ShowFormData()
    End Sub
End Class

