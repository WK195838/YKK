Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO

Partial Class DASW_NewCommission
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
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
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
        Response.Cookies("PGM").Value = "DASW_NewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        'Response.Cookies("pUserID").Value = "it013"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定菜單各程式 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim i As Integer = 0
        Dim SQL As String
       
        '表單
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '006001' And FormNo <= '006099' "
        SQL = SQL + "  And (IniAuthority = '0' "
        SQL = SQL + "       Or (IniAuthority = '1' "
        SQL = SQL + "           And (IniUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
      
        For i = 0 To DBAdapter1.Rows.Count - 1
            '006001
            If DBAdapter1.Rows(i).Item("FormNo") = "006001" Then
                LFun11.Enabled = True
                LFun11.NavigateUrl = "DASW_DISPOSALSheet01.aspx?pFormNo=006001" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID") & _
                                                             "&pUserID=" & Request.QueryString("pUserID")
            End If
            '006002
            If DBAdapter1.Rows(i).Item("FormNo") = "006002" Then
                LFun12.Enabled = True
                LFun12.NavigateUrl = "DASW_NOCopy.aspx?pFormNo=006002" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID") & _
                                                             "&pUserID=" & Request.QueryString("pUserID")
            End If
          
        Next

    End Sub
   
End Class
