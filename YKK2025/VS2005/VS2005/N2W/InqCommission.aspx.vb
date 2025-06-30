Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class InqCommission
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
            If dtForm.Rows(i)("FormNo") = "003101" Then
                LFun01.Enabled = True
                LFun01.NavigateUrl = "ISMSinqCommission.aspx?pFormNo=003101" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If
          
            '交際費申請
            If dtForm.Rows(i)("FormNo") = "003105" Then
                LFun02.Enabled = True
                LFun02.NavigateUrl = "ExpenseinqCommission.aspx?pFormNo=003105" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If

            ' 最終再檢驗處理報告書
            If dtForm.Rows(i)("FormNo") = "003106" Then
                LFun06.Enabled = True
                LFun06.NavigateUrl = "FinalcheckinqCommission.aspx?pFormNo=003106" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If

            ' 最終再檢驗處理報告書-修改
            If dtForm.Rows(i)("FormNo") = "003107" Then
                LFun07.Enabled = True
                LFun07.NavigateUrl = "FinalcheckModinqCommission.aspx?pFormNo=003107" & _
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
    Protected Sub LFun01_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun02_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun02.Init
        LFun02.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun02.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun06_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun06.Init
        LFun06.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun06.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun07_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun07.Init
        LFun07.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun07.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub
End Class

