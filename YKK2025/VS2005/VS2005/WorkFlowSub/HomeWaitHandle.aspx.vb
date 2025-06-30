Imports System.Data
Imports System.Data.OleDb

Partial Class HomeWaitHandle
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HomeWaitHandle.aspx"
        If Not Me.IsPostBack Then
            WaitHandle_DataList()
        End If
    End Sub

    Sub WaitHandle_DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        Dim Http As String = System.Configuration.ConfigurationManager.AppSettings("Http")  'Http Address
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        DHighCount.Text = "(0)"
        DLabel1.Text = ""
        DLabel1.Visible = False
        DLabel2.Text = ""
        DLabel2.Visible = False
        DLabel3.Text = ""
        DLabel3.Visible = False
        DLabel4.Text = ""
        DLabel4.Visible = False
        DLabel5.Text = ""
        DLabel5.Visible = False
        DLabel6.Text = ""
        DLabel6.Visible = False

        OleDbConnection1.Open()
        '待處理件
        SQL = "SELECT "
        SQL = SQL + "Count(*) As Low "
        SQL = SQL + "FROM V_WaitHandle_ApproveHome "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And (Sts = '0'  Or  Sts = '4') "
        SQL = SQL + "  And DecideID = '" + Request.QueryString("pUserID") + "'"
        '
        SQL = SQL + "  And FormNo + Str(FormSno, Len(FormSno)) Not in (select FormNo + str(FormSno, Len(FormSno)) from Q_AgentApprov ) "
        SQL = SQL + "  And FormNo + Str(FormSno, Len(FormSno)) Not in (select FormNo + str(FormSno, Len(FormSno)) from Q_WaitAutoFunding ) "
        '
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "WaitHandle")
        If DBDataSet1.Tables("WaitHandle").Rows(0).Item("Low") > 0 Then
            DHighCount.Text = "(" + DBDataSet1.Tables("WaitHandle").Rows(0).Item("Low").ToString + ")"
            DLabel1.Text = "請馬上處理(連結)"
            DLabel1.NavigateUrl = Http + "/Portal/工作流程/待處理/tabid/90/Default.aspx"
            DLabel1.Visible = True
        End If
        '
        DBDataSet1.Clear()
        '
        '批簽
        'SQL = "SELECT Top 5 FormNo, ProcName, URL "
        SQL = "SELECT Top 4 FormNo, ProcName, URL "
        SQL = SQL & "FROM M_BatchApprovTask "
        SQL = SQL & "Where Active > '0' "
        SQL = SQL & "And UserID = '" + Request.QueryString("pUserID") + "'"
        SQL = SQL & "Order by Unique_ID "
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "BATask")
        For i = 0 To DBDataSet1.Tables("BATask").Rows.Count - 1
            If i = 0 Then
                DLabel2.Text = DBDataSet1.Tables("BATask").Rows(i).Item("ProcName").ToString
                DLabel2.NavigateUrl = DBDataSet1.Tables("BATask").Rows(i).Item("URL").ToString
                DLabel2.Visible = True
            End If
            If i = 1 Then
                DLabel3.Text = DBDataSet1.Tables("BATask").Rows(i).Item("ProcName").ToString
                DLabel3.NavigateUrl = DBDataSet1.Tables("BATask").Rows(i).Item("URL").ToString
                DLabel3.Visible = True
            End If
            If i = 2 Then
                DLabel4.Text = DBDataSet1.Tables("BATask").Rows(i).Item("ProcName").ToString
                DLabel4.NavigateUrl = DBDataSet1.Tables("BATask").Rows(i).Item("URL").ToString
                DLabel4.Visible = True
            End If
            If i = 3 Then
                DLabel5.Text = DBDataSet1.Tables("BATask").Rows(i).Item("ProcName").ToString
                DLabel5.NavigateUrl = DBDataSet1.Tables("BATask").Rows(i).Item("URL").ToString
                DLabel5.Visible = True
            End If
            'If i = 4 Then
            '    DLabel6.Text = DBDataSet1.Tables("BATask").Rows(i).Item("ProcName").ToString
            '    DLabel6.NavigateUrl = DBDataSet1.Tables("BATask").Rows(i).Item("URL").ToString
            '    DLabel6.Visible = True
            'End If
        Next
        '
        OleDbConnection1.Close()
        '
        'COVID-19
        If InStr(UCase(Request.QueryString("pUserID")), "IT") > 0 And _
           UCase(Request.QueryString("pUserID")) <> "IT002" And _
           UCase(Request.QueryString("pUserID")) <> "IT005" And _
           UCase(Request.QueryString("pUserID")) <> "IT014" Then

            DLabel6.Text = "[在家勤務]-刷卡"
            DLabel6.NavigateUrl = "http://10.245.1.6/WorkFlowSub/HomeWorkSign.aspx?puserid=" & Request.QueryString("pUserID")
            DLabel6.Visible = True

        End If
        '
        'If UCase(Request.QueryString("pUserID")) = "GAS002" Or UCase(Request.QueryString("pUserID")) = "GAS009" Then
        '    DLabel2.Text = "經費批簽(連結)"
        '    DLabel2.NavigateUrl = Http + "/portal/管理/經費系統/經費批簽/tabid/444/Default.aspx"
        '    DLabel2.Visible = True
        'End If


    End Sub
End Class
