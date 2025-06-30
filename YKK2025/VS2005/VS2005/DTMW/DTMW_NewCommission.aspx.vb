Imports System.Data
Imports System.Data.OleDb

Partial Class DTMW_NewCommission
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



    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()        '設定共用參數


        If Not Me.IsPostBack Then
            SetMainMenu()         '設定主畫面
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
        BPrint.Attributes.Add("onClick", "RunExe('" + Request.QueryString("pUserID") + " ,')") '列印表單
        BDTMP.Attributes.Add("onClick", "RunExcel()") '優先度

        'BPrint.Attributes.Add("onClick", "RunExe('it013')") '找客戶


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

        SQL = "  SELECT *,case when substring(formname,1,5) = '新色依賴書' then 1 else 2 end as Sno FROM M_Form"
        SQL = SQL + " where formno between '005001' and '005099'"
        SQL = SQL + " and active =1 order by sno,formno"
        DFormNo.Items.Clear()
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("FormName")
            ListItem1.Value = DBAdapter1.Rows(i).Item("FormNo")
            DFormNo.Items.Add(ListItem1)
        Next

    End Sub

    Protected Sub BNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNew.Click
        Dim SQL As String
        Dim wFormno, wURL As String
        wFormno = DFormNo.SelectedValue


        SQL = "  SELECT * FROM  M_referp"
        SQL = SQL + " where dkey  = 'URL-" + wFormno + "'"
        wURL = ""

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            wURL = DBAdapter1.Rows(0).Item("Data")
        End If
        
        'Response.Write(wURL + "?pFormNo=" & DFormNo.SelectedValue & "&pFormSno=0" & _
        '                                                         "&pStep=1" & _
        '                                                         "&pSeqNo=0" & _
        '                                                         "&pApplyID=" & Response.Cookies("UserID").Value & _
        '                                                         "&pUserID=" & Response.Cookies("UserID").Value)
        Dim a As String

        a = wURL + "?pFormNo=" & DFormNo.SelectedValue & "&pFormSno=0" & "&pStep=1" & "&pSeqNo=0" & "&pApplyID=" & Response.Cookies("UserID").Value & "&pUserID=" & Response.Cookies("UserID").Value


        Response.Redirect(wURL + "?pFormNo=" & DFormNo.SelectedValue & "&pFormSno=0" & "&pStep=1" & "&pSeqNo=0" & "&pApplyID=" & Response.Cookies("UserID").Value & "&pUserID=" & Response.Cookies("UserID").Value)


        ' Response.Redirect("MPMProcessesSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
        '  Response.Redirect("HRWYM_TimeOffSheet_02.aspx?pFormNo=001204 " & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")



    End Sub
End Class
