Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class StockNewCommission
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
    Dim wApplyID As String          '申請者ID

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
        Response.Cookies("PGM").Value = "NewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        BStockIn.Attributes.Add("onClick", "RunExcelIN()") '待入庫明細
        BStockOut.Attributes.Add("onClick", "RunExcelOUT()") '待出庫明細
        BSTOCKCHECK.Attributes.Add("onClick", "RunExcelCHECK()") '在庫查詢
        'BPrint.Attributes.Add("onClick", "RunExcelPrint()") '列印表單
        BPrint.Attributes.Add("onClick", "RunExe('" + Request.QueryString("pUserID") + " ,')") '列印表單
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
        SQL = SQL + "  And FormNo >= '003112' And FormNo <= '003113' "
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
            If dtForm.Rows(i)("FormNo") = "003112" Then
                LFun01.Enabled = True

                LFun01.NavigateUrl = "StockInSheet_01.aspx?pFormNo=003112" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")
                LFun03.Enabled = True
                LFun03.NavigateUrl = "StockIninqCommission.aspx?pFormNo=003112" & _
                                                                 "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")

            End If
            If dtForm.Rows(i)("FormNo") = "003113" Then

                LFun02.Enabled = True
                LFun02.NavigateUrl = "StockOutSheet_01.aspx?pFormNo=003113" & _
                                                                  "&pFormSno=0" & _
                                                                  "&pStep=" & xStep & _
                                                                  "&pSeqNo=0" & _
                                                                  "&pUserID=" & Request.QueryString("pUserID") & _
                                                                  "&pApplyID=" & Request.QueryString("pUserID")

                LFun04.Enabled = True
                LFun04.NavigateUrl = "StockOutinqCommission.aspx?pFormNo=003113" & _
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

    Protected Sub LFun03_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun03.Init
        LFun03.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun03.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Protected Sub LFun04_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun04.Init
        LFun04.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun04.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub


End Class

