Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class InqCommissionSales
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

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
        Response.Cookies("PGM").Value = "InqCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
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

        '
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '003101' And FormNo <= '003199' "
        SQL = SQL + "  And (IniAuthority = '0' "
        SQL = SQL + "       Or (IniAuthority = '1' "
        SQL = SQL + "           And (IniUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUsername like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "
        Dim dtForm As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dtForm.Rows.Count - 1
            'ISMS 資訊工作日誌
     
            If dtForm.Rows(i)("FormNo") = "003102" Then
                LFun02.Enabled = True
                LFun02.NavigateUrl = "DiscountinqCommission.aspx?pFormNo=003102" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If dtForm.Rows(i)("FormNo") = "003103" Then
                LFun03.Enabled = True
                LFun03.NavigateUrl = "CustomerInfoinqCommission.aspx?pFormNo=003103" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If dtForm.Rows(i)("FormNo") = "003104" Then
                LFun04.Enabled = True
                LFun04.NavigateUrl = "CustomerInfoModinqCommission.aspx?pFormNo=003104" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If

        Next


    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     菜單各功能滑鼠處理
    '**
    '*****************************************************************
    Protected Sub LFun02_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun02.Init
        LFun02.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun02.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun03_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun03.Init
        LFun03.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun03.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun04_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun04.Init
        LFun04.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun04.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub


End Class

