Imports System.Data
Imports System.Data.OleDb

Partial Class MPM_NewCommission
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
    Dim NowDateTime As String       '現在日期時間
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()        '設定共用參數
        SetMainMenu()         '設定主畫面

        If Not Me.IsPostBack Then
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
        Response.Cookies("PGM").Value = "HRWYM_NewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定菜單各程式 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim SQL As String
        Dim i As Integer
1:
        SQL = "  SELECT * FROM M_Form"
        SQL = SQL + " where formno between '004001' and '004009'"
        SQL = SQL + " and active =1"

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("FormName")
            ListItem1.Value = DBAdapter1.Rows(i).Item("FormNo")
            DFormNo.Items.Add(ListItem1)
        Next

    End Sub

    Protected Sub BNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNew.Click
        If DFormNo.SelectedValue = "004001" Then
            Response.Redirect("MPMProcessesSheet_01.aspx?pFormNo=" & DFormNo.SelectedValue & "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID") & _
                                                             "&pUserID=" & Request.QueryString("pUserID"))

            ' Response.Redirect("MPMProcessesSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
            '  Response.Redirect("HRWYM_TimeOffSheet_02.aspx?pFormNo=001204 " & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")


        End If
       
    End Sub
End Class
