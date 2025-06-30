Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

Partial Class BusinessTax
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject             ' 操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim uWFSCommon As New WFS.CommonService
    Dim NowDateTime As String               ' 現在日時
    Dim xUserID As String                   ' 使用者ID
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        If Not IsPostBack Then
            SetFunctionList("Default", "")              ' 設定可使用之機能

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        uWFSCommon.Timeout = 900 * 1000
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile")
        '隱藏欄位
        xFun.Style("left") = -500 & "px"
        xFunID.Style("left") = -500 & "px"
        xFileName.Style("left") = -500 & "px"
        xFilePath.Style("left") = -500 & "px"
        BINTAX.Visible = False
        BOUTTAX.Visible = False
        BDATA.Visible = False
        BASSETSTAX.Visible = False
        BINTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('TAX')}"
        BASSETSTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('TAX')}"
        BOUTTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('TAX')}"
        BDATA.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('TAX')}"


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetFunctionList)
    '**     設定可使用之機能
    '**
    '*****************************************************************
    Sub SetFunctionList(ByVal pAction As String, ByVal pValue As String)
        Dim sql As String
        Select Case pAction
            Case "Default"
                Dim i As Integer
                '
                DTAX.Items.Clear()
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "003" & "'"
                sql = sql & "   And DKey = '" & xUserID & "'"
                sql = sql & "Order by Data "
                Dim dt_Referp_1 As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Referp_1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dt_Referp_1.Rows(i).Item("Data")
                    ListItem1.Value = dt_Referp_1.Rows(i).Item("Data")
                    DTAX.Items.Add(ListItem1)
                Next
                '

        End Select

        sql = "Select DISTINCT TOP 1 CreateTime FROM M_InputTax  "
        sql = sql & "Order by CreateTime DESC "
        Dim dt_Referp_2 As DataTable = uDataBase.GetDataTable(sql)
        If (dt_Referp_2.Rows.Count = 0) Then
            LINTAX.Text = ""
        Else
            LINTAX.Text = dt_Referp_2.Rows(0).Item("CreateTime")
        End If


        sql = "Select DISTINCT TOP 1 CreateTime FROM M_OutputTax  "
        sql = sql & "Order by CreateTime DESC "
        Dim dt_Referp_3 As DataTable = uDataBase.GetDataTable(sql)
        If (dt_Referp_3.Rows.Count = 0) Then
            LOUTTAX.Text = ""
        Else
            LOUTTAX.Text = dt_Referp_3.Rows(0).Item("CreateTime")
        End If
        sql = "Select DISTINCT TOP 1 CreateTime FROM M_ZeroTax  "
        sql = sql & "Order by CreateTime DESC "
        Dim dt_Referp_4 As DataTable = uDataBase.GetDataTable(sql)
        If (dt_Referp_4.Rows.Count = 0) Then
            LZEROTAX.Text = ""
        Else
            LZEROTAX.Text = dt_Referp_4.Rows(0).Item("CreateTime")
        End If

        sql = "Select DISTINCT TOP 1 CreateTime FROM M_AssetsTax  "
        sql = sql & "Order by CreateTime DESC "
        Dim dt_Referp_5 As DataTable = uDataBase.GetDataTable(sql)
        If (dt_Referp_5.Rows.Count = 0) Then
            LASSETSTAX.Text = ""
        Else
            LASSETSTAX.Text = dt_Referp_5.Rows(0).Item("CreateTime")
        End If


    End Sub
    

    Protected Sub BASSETSTAX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BASSETSTAX.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String

        '
        sql = "Select * From M_Referp "
        sql = sql & " Where DATA = '" & DTAX.SelectedValue & "'"
        sql = sql & " And DKey = '" & xUserID & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "ASSETSTAX"
            xFunID.Text = DTAX.SelectedValue
            '

            SetFunctionList("ASSETSTAX", DTAX.SelectedItem.Text)
            '
            xFileName.Text = DTAX.SelectedValue + ".xlsm"
            xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
            '
            'BASSETSTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否取消此[" + DASSETSTAX.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此功能不存在，請確認!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub

    Protected Sub BOUTTAX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BOUTTAX.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String

        '
        sql = "Select * From M_Referp "
        sql = sql & " Where DATA = '" & DTAX.SelectedValue & "'"
        sql = sql & " And DKey = '" & xUserID & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "OUTTAX"
            xFunID.Text = DTAX.SelectedValue
            '

            SetFunctionList("OUTTAX", DTAX.SelectedItem.Text)
            '
            xFileName.Text = DTAX.SelectedValue + ".xlsm"
            xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
            '
            'BASSETSTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否取消此[" + DASSETSTAX.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此功能不存在，請確認!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub

    Protected Sub BDATA_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDATA.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String

        '
        sql = "Select * From M_Referp "
        sql = sql & " Where DATA = '" & DTAX.SelectedValue & "'"
        sql = sql & " And DKey = '" & xUserID & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "DATATAX"
            xFunID.Text = DTAX.SelectedValue
            '

            SetFunctionList("DATATAX", DTAX.SelectedItem.Text)
            '
            xFileName.Text = DTAX.SelectedValue + ".xlsm"
            xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
            '
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此功能不存在，請確認!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub


    Protected Sub BINTAX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BINTAX.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String

        '
        sql = "Select * From M_Referp "
        sql = sql & " Where DATA = '" & DTAX.SelectedValue & "'"
        sql = sql & " And DKey = '" & xUserID & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "INTAX"
            xFunID.Text = DTAX.SelectedValue
            '

            SetFunctionList("INTAX", DTAX.SelectedItem.Text)
            '
            xFileName.Text = DTAX.SelectedValue + ".xls"
            xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
            '
            'BASSETSTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否取消此[" + DASSETSTAX.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此功能不存在，請確認!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub

    Protected Sub BTAX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BTAX.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String
        '
        sql = "Select * From M_Referp "
        sql = sql & " Where DATA = '" & DTAX.SelectedValue & "'"
        sql = sql & " And DKey = '" & xUserID & "'"
        'MsgBox(sql)
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "TAX"
            xFunID.Text = DTAX.SelectedValue
            '
            SetFunctionList("TAX", DTAX.SelectedItem.Text)
            '
            If (DTAX.SelectedValue = "進項稅額資料查詢") Then
                xFileName.Text = DTAX.SelectedValue + "_" + xUserID + ".xlsm"
            Else
                xFileName.Text = DTAX.SelectedValue + ".xlsm"
            End If
            xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
            Select Case DTAX.SelectedValue
                Case "進項稅額資料檢核"
                    BINTAX.Visible = True
                Case "進項稅額資料查詢"
                    BINTAX.Visible = True
                Case "銷項稅額資料檢核"
                    BOUTTAX.Visible = True
                Case "固定資產退稅資料檢核"
                    BASSETSTAX.Visible = True
                Case "固定資產退稅資料查詢"
                    BASSETSTAX.Visible = True
                Case "稅碼對照表"
                    BINTAX.Visible = True
                Case "資料匯出"
                    BDATA.Visible = True
            End Select

            '
            'BINTAX.Attributes("onclick") = "var ok=window.confirm('" + "是否取消此[" + DINTAX.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"

        Else
            errcode = 9001
        End If
            '
            If errcode <> 0 Then
                If errcode = 9001 Then msg = "此功能不存在，請確認!"
                uJavaScript.PopMsg(Me, msg)
            End If
    End Sub
End Class
