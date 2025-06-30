Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class OpenPOPicker
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim wBuyer As String            'Parameter=pBuyer
    Dim wFun As String              'Parameter=pFun

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        '
        If Not Me.IsPostBack Then   '不是PostBack
            DKey.Text = ""
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "OpenPOPicker.aspx"                                         '程式名
        wBuyer = Request.QueryString("pBuyer")
        wFun = Request.QueryString("pFun")
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BKey_Click)
    '**     搜尋按鈕按下時
    '**
    '*****************************************************************
    Protected Sub BKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BKey.Click
        If DKey.Text <> "" Then
            DataList()
        Else
            uJavaScript.PopMsg(Me, "無輸入資料無法搜尋,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     開發資料列表
    '**
    '*****************************************************************
    Protected Sub DataList()
        Dim SQL As String
        Dim xBuyerGroup As String = ""
        '
        SQL = "Select BuyerGroup "
        SQL &= "From M_ControlRecord "
        SQL &= "Where Buyer = '" & wBuyer & "' "
        Dim dtControlRecord As DataTable = uDataBase.GetDataTable(SQL)
        If dtControlRecord.Rows.Count > 0 Then
            xBuyerGroup = dtControlRecord.Rows(0).Item("BuyerGroup")
        End If
        '
        If wFun = "PROGRESS" Then
            SQL = "Select Top 30 PO "
            SQL &= "From B_OrderProgress "
            SQL &= "Where PO Like '%" & DKey.Text & "%' "
            SQL &= "  And CustomerBuyer = '" & wBuyer & "' "
            SQL &= "Group by PO "
            SQL &= "Order by PO Desc "
        Else
            SQL = "Select Top 30 PO "
            SQL &= "From B_PackingInstruction "
            SQL &= "Where PO Like '%" & DKey.Text & "%' "
            SQL &= "  And CustomerBuyer = '" & wBuyer & "' "
            SQL &= "Group by PO "
            SQL &= "Order by PO Desc "
        End If

        '---------------------------------------------------------
        'If fpObj.GetFunctionCode(xBuyerGroup, 2) = "P" Then
        '    SQL = "Select Top 30 BM1 As PO "
        '    SQL &= "From B_CustomerRequest "
        '    SQL &= "Where BM1 Like '%" & DKey.Text & "%' "
        '    SQL &= "  And Buyer = '" & wBuyer & "' "
        '    SQL &= "Group by BM1 "
        '    SQL &= "Order by BM1 Desc "
        'Else
        '    SQL = "Select Top 30 E1 As PO "
        '    SQL &= "From B_CustomerRequest "
        '    SQL &= "Where E1 Like '%" & DKey.Text & "%' "
        '    SQL &= "  And Buyer = '" & wBuyer & "' "
        '    SQL &= "Group by E1 "
        '    SQL &= "Order by E1 Desc "
        'End If
        '
        Dim dtCustomerRequest As DataTable = uDataBase.GetDataTable(SQL)
        If dtCustomerRequest.Rows.Count > 0 Then
            GridView1.DataSource = dtCustomerRequest
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_SelectedIndexChanged)
    '**     取得所選擇開發資料
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim xPO As String = LTrim(RTrim(GridView1.SelectedRow.Cells(1).Text))
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "form1.DPO", xPO) + _
                  "window.close();" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
End Class
