Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class NewCommission
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
        Response.Cookies("PGM").Value = "NewCommission.aspx"
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

        'Dim SQL As String = "Select DivID FROM M_Users "
        'SQL = SQL & " Where Active =  '1' "
        'SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
        'Dim dt_Users As DataTable = uDataBase.GetDataTable(SQL)
        'If dt_Users.Rows.Count > 0 Then
        '    If dt_Users.Rows(0)("DivID") = "1108213" Then xStep = 2
        'End If
        If Request.QueryString("pUserID") = "mk028" Or Request.QueryString("pUserID") = "mk019" Or Request.QueryString("pUserID") = "mk035" Then xStep = 2
        '
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '001151' And FormNo <= '001199' "
        SQL = SQL + "  And (IniAuthority = '0' "
        SQL = SQL + "       Or (IniAuthority = '1' "
        SQL = SQL + "           And (IniUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "
        Dim dtForm As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dtForm.Rows.Count - 1
            '營業用登錄申請單
            If dtForm.Rows(i)("FormNo") = "001151" Then
                LFun01.Enabled = True
                'ISOS-IRW
                LFun01.NavigateUrl = "http://10.245.0.3/ISOS/ps_menu.aspx?pUserID=" & Request.QueryString("pUserID")

                'LFun01.NavigateUrl = "ItemRegisterSheet_01.aspx?pFormNo=001151" & _
                '                                                  "&pFormSno=0" & _
                '                                                  "&pStep=" & xStep & _
                '                                                  "&pSeqNo=0" & _
                '                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                '                                                  "&pApplyID=" & Request.QueryString("pUserID")
            End If
            'ZIP用登錄申請單
            If dtForm.Rows(i)("FormNo") = "001152" Then
                LFun11.Enabled = True
                LFun11.NavigateUrl = "ItemRegisterZIPSheet_01.aspx?pFormNo=001152" & _
                                                                 "&pFormSno=0" & _
                                                                 "&pStep=1" & _
                                                                 "&pSeqNo=0" & _
                                                                 "&pUserID=" & Request.QueryString("pUserID") & _
                                                                 "&pApplyID=" & Request.QueryString("pUserID")
            End If
            'SLD用登錄申請單
            If dtForm.Rows(i)("FormNo") = "001153" Then
                LFun12.Enabled = True
                LFun12.NavigateUrl = "ItemRegisterSLDSheet_01.aspx?pFormNo=001153" & _
                                                                 "&pFormSno=0" & _
                                                                 "&pStep=1" & _
                                                                 "&pSeqNo=0" & _
                                                                 "&pUserID=" & Request.QueryString("pUserID") & _
                                                                 "&pApplyID=" & Request.QueryString("pUserID")
            End If
            'CH用登錄申請單
            If dtForm.Rows(i)("FormNo") = "001154" Then
                LFun13.Enabled = True
                LFun13.NavigateUrl = "ItemRegisterCHSheet_01.aspx?pFormNo=001154" & _
                                                                "&pFormSno=0" & _
                                                                "&pStep=1" & _
                                                                "&pSeqNo=0" & _
                                                                "&pUserID=" & Request.QueryString("pUserID") & _
                                                                "&pApplyID=" & Request.QueryString("pUserID")
            End If
            'SLD登錄申請單(工廠)
            If dtForm.Rows(i)("FormNo") = "001155" Then
                LFun02.Enabled = True
                LFun02.NavigateUrl = "ItemRegisterFSLDSheet_01.aspx?pFormNo=001155" & _
                                                                 "&pFormSno=0" & _
                                                                 "&pStep=1" & _
                                                                 "&pSeqNo=0" & _
                                                                 "&pUserID=" & Request.QueryString("pUserID") & _
                                                                 "&pApplyID=" & Request.QueryString("pUserID")
            End If
            '21. 單價承認申請書
            If dtForm.Rows(i)("FormNo") = "001161" Then
                LFun21.Enabled = True
                LFun21.NavigateUrl = "PriceInforSheet_01.aspx?pFormNo=001161" & _
                                                                 "&pFormSno=0" & _
                                                                 "&pStep=1" & _
                                                                 "&pSeqNo=0" & _
                                                                 "&pUserID=" & Request.QueryString("pUserID") & _
                                                                 "&pApplyID=" & Request.QueryString("pUserID")
            End If
            '31. ITEM未受注報告書
            If dtForm.Rows(i)("FormNo") = "001171" Then
                LFun31.Enabled = True
                LFun31.NavigateUrl = "IRWNoOrderReportSheet_01.aspx?pFormNo=001171" & _
                                                                 "&pFormSno=0" & _
                                                                 "&pStep=1" & _
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
    Protected Sub LFun11_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun11.Init
        LFun11.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun11.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub
    Protected Sub LFun12_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun12.Init
        LFun12.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun12.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub
    Protected Sub LFun13_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun13.Init
        LFun13.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun13.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub
    Protected Sub LFun21_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun21.Init
        LFun21.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun21.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun31_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun31.Init
        LFun31.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun31.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub
End Class

