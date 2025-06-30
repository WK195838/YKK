Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

Partial Class InfDSJobList
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
    Dim NowDateTime As String               ' 現在日時
    Dim xUserID As String                   ' 使用者ID
    Dim xBuyer As String                    ' 客戶代碼
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        If Not IsPostBack Then
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetJobList("Default", "")                   ' 設定可使用功能
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        xBuyer = UCase(Request.QueryString("pBuyer"))
        DFileName.Style("left") = -500 & "px"
        DFilePath.Style("left") = -500 & "px"
        DFileName1.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        '動作按鈕設定
        BJob.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此工作？" + "');if(!ok){return false;}"
        BExcel.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟Excel準備資料？" + "');if(!ok){return false;} else {CheckAttribute('Excel')}"
        BRun.Attributes("onclick") = "var ok=window.confirm('" + "是否執行作業？" + "');if(!ok){return false;} else {CheckAttribute('Run')}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetCustomerBuyer)
    '**     設定可使用之客戶
    '**
    '*****************************************************************
    Sub SetJobList(ByVal pAction As String, ByVal pValue As String)
        Dim sql As String
        Select Case pAction
            Case "Default"
                Dim i As Integer
                '
                DJob.Items.Clear()
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "020" & "'"
                sql = sql & "   And DKey = '" & xBuyer & "'"
                sql = sql & "Order by Data "
                Dim dt_Referp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Referp.Rows.Count - 1
                    '
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Mid(dt_Referp.Rows(i).Item("Data"), 1, InStr(dt_Referp.Rows(i).Item("Data"), "/") - 1)
                    ListItem1.Value = Mid(dt_Referp.Rows(i).Item("Data"), InStr(dt_Referp.Rows(i).Item("Data"), "/") + 1)
                    DJob.Items.Add(ListItem1)
                Next
            Case Else
                DJob.Items.Clear()
                '
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "020" & "'"
                sql = sql & "   And DKey = '" & xBuyer & "'"
                sql = sql & "   And Data Like '%" & pValue & "%'"
                sql = sql & "Order by Data "
                Dim dt_Referp As DataTable = uDataBase.GetDataTable(sql)
                If dt_Referp.Rows.Count > 0 Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Mid(dt_Referp.Rows(0).Item("Data"), 1, InStr(dt_Referp.Rows(0).Item("Data"), "/") - 1)
                    ListItem1.Value = Mid(dt_Referp.Rows(0).Item("Data"), InStr(dt_Referp.Rows(0).Item("Data"), "/") + 1)
                    DJob.Items.Add(ListItem1)
                End If
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskStatus)
    '**     設定各動作執行結果狀態
    '**     pAction=動作名稱    pLeft=顯示(>0)或不顯示(<0)  pStatus=顯示圖像
    '**
    '*****************************************************************
    Sub SetTaskStatus(ByVal pAction As String, ByVal pLeft As Integer, ByVal pStatus As Integer)
        Select Case pAction
            Case "StsJob"
                StsJob.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsJob.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsJob.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsExcel"
                StsExcel.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsExcel.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsExcel.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsRun"
                StsRun.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsRun.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsRun.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case Else
                StsJob.Style("left") = pLeft & "px"
                StsExcel.Style("left") = pLeft & "px"
                StsRun.Style("left") = pLeft & "px"
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskProgress)
    '**     設定各動作執行中Task Bar
    '**     pAction=動作名稱    pShow=顯示(>0)或不顯示(<0)
    '**
    '*****************************************************************
    Sub SetTaskProgress(ByVal pAction As String, ByVal pLeft As Integer)
        Select Case pAction
            Case "ProExcel"
                ProExcel.Style("left") = pLeft & "px"
            Case "ProRun"
                ProRun.Style("left") = pLeft & "px"
            Case Else
                ProExcel.Style("left") = pLeft & "px"
                ProRun.Style("left") = pLeft & "px"
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskButton)
    '**     設定各動作按鈕
    '**     pAction=動作名稱    pShow=顯示或不顯示
    '**
    '*****************************************************************
    Sub SetTaskButton(ByVal pAction As String, ByVal pShow As Boolean)
        Select Case pAction
            Case "BReset"
                BReset.Visible = pShow
            Case "BJob"
                BJob.Visible = pShow
            Case "BExcel"
                BExcel.Visible = pShow
            Case "BRun"
                BRun.Visible = pShow
            Case Else
                BReset.Visible = False
                BJob.Visible = True
                BExcel.Visible = False
                BRun.Visible = False
                '
                'Test-Start
                'BReset.Visible = True
                'BJob.Visible = True
                'BExcel.Visible = True
                'BRun.Visible = True
                'Test-End
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BJob_Click)
    '**     工作選擇
    '**
    '*****************************************************************
    Protected Sub BJob_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BJob.Click
        '
        SetTaskButton("BJob", False)
        SetTaskButton("BReset", True)
        SetTaskButton("BExcel", True)
        SetTaskStatus("StsJob", 601, 0)
        SetJobList("Job", DJob.SelectedValue)
        '
        DFileName.Text = DJob.SelectedValue + "_DataPrepare.xlsm"
        DFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
        '
        DFileName1.Text = DJob.SelectedValue + ".exe"
        DFilePath1.Text = uCommon.GetAppSetting("ProgramFile") + DFileName1.Text
        '
        BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消工作[" + DJob.SelectedItem.Text + "]，" + "並且重新選擇工作？" + "');if(!ok){return false;}"
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BReset_Click)
    '**     工作重新選擇
    '**
    '*****************************************************************
    Protected Sub BReset_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReset.Click
        '
        SetTaskButton("Default", False)             ' 動作按鈕圖像
        SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
        SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
        SetJobList("Default", "")                   ' 設定可使用工作
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BExcel_Click)
    '**     準備客戶資料
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        SetTaskProgress("ProExcel", -500)
        SetTaskButton("BExcel", False)
        SetTaskButton("BReset", True)
        SetTaskButton("BRun", True)
        SetTaskStatus("StsExcel", 601, 0)
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BRun_Click)
    '**     執行Program 
    '**
    '*****************************************************************
    Protected Sub BRun_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRun.Click
        '
        SetTaskProgress("ProRun", -500)
        SetTaskButton("BRun", False)
        SetTaskButton("BReset", True)
        SetTaskStatus("StsRBS", 601, 0)
        '
    End Sub

End Class
