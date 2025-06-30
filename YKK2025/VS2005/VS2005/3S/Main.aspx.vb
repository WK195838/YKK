Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

Partial Class Main
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
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
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
        '動作按鈕設定
        BJGSFun.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此功能？" + "');if(!ok){return false;}"
        BRISFun.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此功能？" + "');if(!ok){return false;}"

        BJGS.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟準備資料？" + "');if(!ok){return false;} else {CheckAttribute('JGS')}"
        BRIS.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟準備資料？" + "');if(!ok){return false;} else {CheckAttribute('RIS')}"
        BBIS.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('BIS')}"
        BDGS.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('DGS')}"
        BRBS.Attributes("onclick") = "var ok=window.confirm('" + "是否執行？" + "');if(!ok){return false;} else {CheckAttribute('RBS')}"
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
            Case "BJGSFun"
                BJGSFun.Visible = pShow
            Case "BRISFun"
                BRISFun.Visible = pShow
            Case "BJGS"
                BJGS.Visible = pShow
            Case "BRIS"
                BRIS.Visible = pShow
            Case "BBIS"
                BBIS.Visible = pShow
            Case "BDGS"
                BDGS.Visible = pShow
            Case "BRBS"
                BRBS.Visible = pShow
            Case Else
                BReset.Visible = False
                BJGSFun.Visible = True
                BRISFun.Visible = True
                BJGS.Visible = False
                BRIS.Visible = False
                BBIS.Visible = False
                BDGS.Visible = False
                BRBS.Visible = False
                '
                'Test-Start
                'BReset.Visible = True
                'BJGSFun.Visible = True
                'BRISFun.Visible = True
                'BJGS.Visible = True
                'BRIS.Visible = True
                'BBIS.Visible = True
                'BDGS.Visible = True
                'BRBS.Visible = True
                'Test-End
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
            Case "StsJGSFun"
                StsJGSFun.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsJGSFun.ImageUrl = "iMages/OK.png"
                    Else
                        StsJGSFun.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case "StsRISFun"
                StsRISFun.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsRISFun.ImageUrl = "iMages/OK.png"
                    Else
                        StsRISFun.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case "StsJGS"
                StsJGS.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsJGS.ImageUrl = "iMages/OK.png"
                    Else
                        StsJGS.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case "StsRIS"
                StsRIS.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsRIS.ImageUrl = "iMages/OK.png"
                    Else
                        StsRIS.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case "StsBIS"
                StsBIS.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBIS.ImageUrl = "iMages/OK.png"
                    Else
                        StsBIS.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case "StsDGS"
                StsDGS.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsDGS.ImageUrl = "iMages/OK.png"
                    Else
                        StsDGS.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case "StsRBS"
                StsRBS.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsRBS.ImageUrl = "iMages/OK.png"
                    Else
                        StsRBS.ImageUrl = "iMages/NG.png"
                    End If
                End If
            Case Else
                StsJGSFun.Style("left") = pLeft & "px"
                StsRISFun.Style("left") = pLeft & "px"
                StsJGS.Style("left") = pLeft & "px"
                StsRIS.Style("left") = pLeft & "px"
                StsBIS.Style("left") = pLeft & "px"
                StsDGS.Style("left") = pLeft & "px"
                StsRBS.Style("left") = pLeft & "px"
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
            Case "ProJGSFun"
                ProJGSFun.Style("left") = pLeft & "px"
            Case "ProRISFun"
                ProRISFun.Style("left") = pLeft & "px"
            Case "ProJGS"
                ProJGS.Style("left") = pLeft & "px"
            Case "ProRIS"
                ProRIS.Style("left") = pLeft & "px"
            Case "ProBIS"
                ProBIS.Style("left") = pLeft & "px"
            Case "ProDGS"
                ProDGS.Style("left") = pLeft & "px"
            Case "ProRBS"
                ProRBS.Style("left") = pLeft & "px"
            Case Else
                ProJGSFun.Style("left") = pLeft & "px"
                ProRISFun.Style("left") = pLeft & "px"
                ProJGS.Style("left") = pLeft & "px"
                ProRIS.Style("left") = pLeft & "px"
                ProBIS.Style("left") = pLeft & "px"
                ProDGS.Style("left") = pLeft & "px"
                ProRBS.Style("left") = pLeft & "px"
        End Select
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
                DJGSFun.Items.Clear()
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "001" & "'"
                sql = sql & "   And DKey = '" & xUserID & "'"
                sql = sql & "Order by Data "
                Dim dt_Referp_1 As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Referp_1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Mid(dt_Referp_1.Rows(i).Item("Data"), 1, InStr(dt_Referp_1.Rows(i).Item("Data"), "/") - 1)
                    ListItem1.Value = Mid(dt_Referp_1.Rows(i).Item("Data"), InStr(dt_Referp_1.Rows(i).Item("Data"), "/") + 1)
                    DJGSFun.Items.Add(ListItem1)
                Next
                '
                DRISFun.Items.Clear()
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "002" & "'"
                sql = sql & "   And DKey = '" & xUserID & "'"
                sql = sql & "Order by Data "
                Dim dt_Referp_2 As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Referp_2.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Mid(dt_Referp_2.Rows(i).Item("Data"), 1, InStr(dt_Referp_2.Rows(i).Item("Data"), "/") - 1)
                    ListItem1.Value = Mid(dt_Referp_2.Rows(i).Item("Data"), InStr(dt_Referp_2.Rows(i).Item("Data"), "/") + 1)
                    DRISFun.Items.Add(ListItem1)
                Next

            Case Else
                DJGSFun.Items.Clear()
                DRISFun.Items.Clear()
                '
                Dim ListItem1 As New ListItem
                ListItem1.Text = pValue
                ListItem1.Value = xFunID.Text
                If xFun.Text = "JGS" Then
                    DJGSFun.Items.Add(ListItem1)
                Else
                    DRISFun.Items.Add(ListItem1)
                End If
        End Select
    End Sub
    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([500]-UpdateControlRecord)
    '**     更新客戶控制檔-Action Flag
    '**     pAction=動作名稱    pShow=顯示或不顯示
    '***********************************************************************************************
    Public Function UpdateControlRecord(ByVal pID As String, _
                                        ByVal pUserID As String, _
                                        ByVal pAction1 As String, _
                                        ByVal pStatus1 As Integer, _
                                        ByVal pAction2 As String, _
                                        ByVal pStatus2 As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Try
            Dim sql As String
            Select Case pAction1
                Case "Start"
                    If xFun.Text = "JGS" Then
                        sql = "Update M_ControlRecord Set "
                        sql = sql + " Active = '1', "
                        sql = sql + " JGSFun = '2', "
                        sql = sql + " RISFun = '0', "
                        sql = sql + " JGS = '1', "
                        sql = sql + " RIS = '0', "
                        sql = sql + " BIS = '0', "
                        sql = sql + " DGS = '0', "
                        sql = sql + " RBS = '0', "
                        sql = sql + " ModifyUser = '" & pUserID & "', "
                        sql = sql + " ModifyTime = '" & NowDateTime & "' "
                        sql = sql + " Where ID = '" & pID & "' "
                    Else
                        sql = "Update M_ControlRecord Set "
                        sql = sql + " Active = '1', "
                        sql = sql + " JGSFun = '0', "
                        sql = sql + " RISFun = '2', "
                        sql = sql + " JGS = '0', "
                        sql = sql + " RIS = '1', "
                        sql = sql + " BIS = '0', "
                        sql = sql + " DGS = '0', "
                        sql = sql + " RBS = '0', "
                        sql = sql + " ModifyUser = '" & pUserID & "', "
                        sql = sql + " ModifyTime = '" & NowDateTime & "' "
                        sql = sql + " Where ID = '" & pID & "' "
                    End If
                    uDataBase.ExecuteNonQuery(sql)
                Case "End"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where ID = '" & pID & "' "
                    uDataBase.ExecuteNonQuery(sql)
                Case "Reset"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + " Active = '0', "
                    sql = sql + " JGSFun = '0', "
                    sql = sql + " RISFun = '0', "
                    sql = sql + " JGS = '0', "
                    sql = sql + " RIS = '0', "
                    sql = sql + " BIS = '0', "
                    sql = sql + " DGS = '0', "
                    sql = sql + " RBS = '0', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where ID = '" & pID & "' "
                    uDataBase.ExecuteNonQuery(sql)
                Case "Action2"
                    sql = "Update M_ControlRecord Set "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where ID = '" & pID & "' "
                    uDataBase.ExecuteNonQuery(sql)
                Case Else
                    sql = "Update M_ControlRecord Set "
                    sql = sql + pAction1 + " = '" & CStr(pStatus1) & "', "
                    sql = sql + pAction2 + " = '" & CStr(pStatus2) & "', "
                    sql = sql + " ModifyUser = '" & pUserID & "', "
                    sql = sql + " ModifyTime = '" & NowDateTime & "' "
                    sql = sql + " Where ID = '" & pID & "' "
                    uDataBase.ExecuteNonQuery(sql)
            End Select
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([510]-CheckControlRecord)
    '**     檢查客戶控制檔-Action Flag--判斷是否完成作業
    '***********************************************************************************************
    Public Function CheckControlRecord(ByVal pID As String, ByVal pAction As String, ByVal pValue As Integer) As Integer
        Dim RtnCode As Integer = 0
        '
        Dim sql As String
        Try
            sql = "Select * From M_ControlRecord "
            sql = sql & " Where ID = '" & pID & "'"
            Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                If dt_ControlRecord.Rows(0).Item(pAction) <> pValue Then
                    RtnCode = 1
                End If
            Else
                RtnCode = 1
            End If
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BReset_Click)
    '**     客戶重新選擇
    '**
    '*****************************************************************
    Protected Sub BReset_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReset.Click
        If UpdateControlRecord(xFunID.Text, xUserID, "Reset", 0, "Reset", 0) = 0 Then
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetFunctionList("Default", "")              ' 設定可使用之機能
        Else
            uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag)，請連絡系統人員!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BJGSFun_Click)
    '**     選擇JGS機能 
    '**
    '*****************************************************************
    Protected Sub BJGSFun_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BJGSFun.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String
        '
        sql = "Select * From M_ControlRecord "
        sql = sql & " Where ID = '" & DJGSFun.SelectedValue & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "JGS"
            xFunID.Text = DJGSFun.SelectedValue
            mUserID = UCase(dt_ControlRecord.Rows(0).Item("ModifyUser"))
            '
            If dt_ControlRecord.Rows(0).Item("Active") = 0 Then
                If UpdateControlRecord(DJGSFun.SelectedValue, xUserID, "Start", 0, "Start", 0) = 0 Then
                    SetTaskButton("BJGSFun", False)
                    SetTaskButton("BRISFun", False)
                    SetTaskButton("BReset", True)
                    SetTaskButton("BJGS", True)
                    SetTaskStatus("StsJGSFun", 174, 0)
                    SetFunctionList("JGSFun", DJGSFun.SelectedItem.Text)
                    '
                    xFileName.Text = DJGSFun.SelectedValue + "_DataPrepare.xls"
                    xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
                    '
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消此[" + DJGSFun.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
                Else
                    errcode = 9003
                End If
            Else
                If mUserID = xUserID Then
                    errcode = 9004      '同上次使用者(上次不小心)
                    SetTaskButton("BReset", True)
                    SetFunctionList("JGSFun", DJGSFun.SelectedItem.Text)
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消[" + DJGSFun.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
                Else
                    errcode = 9002      '不同上次使用者
                End If
            End If
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此功能不存在，請確認!"
            If errcode = 9002 Then msg = "此功能正由 [" + uWFSCommon.GetUserName(mUserID) + "] 使用中，請確認!"
            If errcode = 9003 Then msg = "無法更新功能控制檔(Flag)，請連絡系統人員!"
            If errcode = 9004 Then msg = "懷疑您上次使用有不正常關閉系統，造成此功能正在使用中! 請[Reset]一次再使用!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BRISFun_Click)
    '**     選擇RIS機能 
    '**
    '*****************************************************************
    Protected Sub BRISFun_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRISFun.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String

        '
        sql = "Select * From M_ControlRecord "
        sql = sql & " Where ID = '" & DRISFun.SelectedValue & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            xFun.Text = "RIS"
            xFunID.Text = DRISFun.SelectedValue
            mUserID = UCase(dt_ControlRecord.Rows(0).Item("ModifyUser"))
            '
            If dt_ControlRecord.Rows(0).Item("Active") = 0 Then
                If UpdateControlRecord(DRISFun.SelectedValue, xUserID, "Start", 0, "Start", 0) = 0 Then
                    SetTaskButton("BJGSFun", False)
                    SetTaskButton("BRISFun", False)
                    SetTaskButton("BReset", True)
                    SetTaskButton("BRIS", True)
                    SetTaskStatus("StsRISFun", 174, 0)
                    SetFunctionList("RISFun", DRISFun.SelectedItem.Text)
                    '
                    xFileName.Text = DRISFun.SelectedValue + "_DataPrepare.xls"
                    xFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + xFileName.Text
                    '
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消此[" + DRISFun.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
                Else
                    errcode = 9003
                End If
            Else
                If mUserID = xUserID Then
                    errcode = 9004      '同上次使用者(上次不小心)
                    SetTaskButton("BReset", True)
                    SetFunctionList("RISFun", DRISFun.SelectedItem.Text)
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消[" + DRISFun.SelectedItem.Text + "]，" + "並且重新選擇功能？" + "');if(!ok){return false;}"
                Else
                    errcode = 9002      '不同上次使用者
                End If
            End If
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此功能不存在，請確認!"
            If errcode = 9002 Then msg = "此功能正由 [" + uWFSCommon.GetUserName(mUserID) + "] 使用中，請確認!"
            If errcode = 9003 Then msg = "無法更新功能控制檔(Flag)，請連絡系統人員!"
            If errcode = 9004 Then msg = "懷疑您上次使用有不正常關閉系統，造成此功能正在使用中! 請[Reset]一次再使用!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BJGS_Click)
    '**     執行JGS-準備資料 
    '**
    '*****************************************************************
    Protected Sub BJGS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BJGS.Click
        '
        '判斷是否完成作業
        Dim wType As String = xFun.Text

        UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, wType, 1)
        '
        Dim Code As Integer
        Code = CheckControlRecord(xFunID.Text, wType, 2)
        If Code <> 0 Then
            Code = CheckControlRecord(xFunID.Text, wType, 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = CheckControlRecord(xFunID.Text, wType, 2)
            If Code <> 0 Then
                Code = CheckControlRecord(xFunID.Text, wType, 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If CheckControlRecord(xFunID.Text, wType, 2) = 0 Then
            '
            Dim sql As String
            sql = "Select * From M_ControlRecord "
            sql = sql & " Where ID = '" & xFunID.Text & "'"
            Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                xFileName.Text = dt_ControlRecord.Rows(0).Item("Program")
            End If
            xFilePath.Text = uCommon.GetAppSetting("ProgramFile") + xFileName.Text
            '
            SetTaskProgress("ProJGS", -500)
            SetTaskButton("BJGS", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BBIS", True)
            UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, "BIS", 1)
            SetTaskStatus("StsJGS", 536, 0)
        Else
            SetTaskStatus("StsJGS", 578, 1)
            uJavaScript.PopMsg(Me, "資料準備未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BRIS_Click)
    '**     執行RIS-準備資料 
    '**
    '*****************************************************************
    Protected Sub BRIS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRIS.Click
        '
        '判斷是否完成作業
        Dim wType As String = xFun.Text

        UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, wType, 1)
            '
        Dim Code As Integer
        Code = CheckControlRecord(xFunID.Text, wType, 2)
        If Code <> 0 Then
            Code = CheckControlRecord(xFunID.Text, wType, 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = CheckControlRecord(xFunID.Text, wType, 2)
            If Code <> 0 Then
                Code = CheckControlRecord(xFunID.Text, wType, 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If CheckControlRecord(xFunID.Text, wType, 2) = 0 Then
            SetTaskProgress("ProRIS", -500)
            SetTaskButton("BRIS", False)
            SetTaskButton("BReset", True)
            '
            Dim sql As String
            sql = "Select * From M_ControlRecord "
            sql = sql & " Where ID = '" & xFunID.Text & "'"
            Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
            If dt_ControlRecord.Rows.Count > 0 Then
                ' Modify-Start 2013/6/10
                If dt_ControlRecord.Rows(0).Item("TableName") <> "" Then
                    xFileName.Text = dt_ControlRecord.Rows(0).Item("Program") + " " + xUserID + " " + dt_ControlRecord.Rows(0).Item("TableName")
                Else
                    xFileName.Text = dt_ControlRecord.Rows(0).Item("Program")
                End If
                'If CInt(Mid(xFunID.Text, 5)) >= 901 Then
                '    ' Master Program
                '    xFileName.Text = dt_ControlRecord.Rows(0).Item("Program") + " " + xUserID + " " + dt_ControlRecord.Rows(0).Item("TableName")
                'Else
                '    xFileName.Text = dt_ControlRecord.Rows(0).Item("Program")
                'End If
                ' Modify-End

                xFilePath.Text = uCommon.GetAppSetting("ProgramFile") + xFileName.Text
                '
                If dt_ControlRecord.Rows(0).Item("Type") = "DGS" Then
                    SetTaskButton("BDGS", True)
                    UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, "DGS", 1)
                ElseIf dt_ControlRecord.Rows(0).Item("Type") = "RBS" Then
                    SetTaskButton("BRBS", True)
                    UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, "RBS", 1)
                Else
                    UpdateControlRecord(xFunID.Text, xUserID, "End", 0, "End", 0)
                End If

                'Else
                '    SetTaskButton("BRBS", True)
                '    UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, "RBS", 1)
                'End If

            End If
            '
            SetTaskStatus("StsRIS", 401, 0)
        Else
            SetTaskStatus("StsRIS", 443, 1)
            uJavaScript.PopMsg(Me, "資料準備未完成，請確認!")
        End If

        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBIS_Click)
    '**     執行BIS Program 
    '**
    '*****************************************************************
    Protected Sub BBIS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBIS.Click
        '
        '判斷是否完成作業
        Dim wType As String = "BIS"

        UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, wType, 1)
        '
        Dim Code As Integer
        Code = CheckControlRecord(xFunID.Text, wType, 2)
        If Code <> 0 Then
            Code = CheckControlRecord(xFunID.Text, wType, 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = CheckControlRecord(xFunID.Text, wType, 2)
            If Code <> 0 Then
                Code = CheckControlRecord(xFunID.Text, wType, 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If CheckControlRecord(xFunID.Text, wType, 2) = 0 Then
            SetTaskProgress("ProBIS", -500)
            SetTaskButton("BBIS", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsBIS", 848, 0)
            UpdateControlRecord(xFunID.Text, xUserID, "End", 0, "End", 0)
        Else
            SetTaskStatus("StsBIS", 890, 1)
            uJavaScript.PopMsg(Me, "程式未執行完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BDGS_Click)
    '**     執行DGS Program 
    '**
    '*****************************************************************
    Protected Sub BDGS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDGS.Click
        '
        '判斷是否完成作業
        Dim wType As String = "DGS"

        UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, wType, 1)
        '
        Dim Code As Integer
        Code = CheckControlRecord(xFunID.Text, wType, 2)
        If Code <> 0 Then
            Code = CheckControlRecord(xFunID.Text, wType, 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = CheckControlRecord(xFunID.Text, wType, 2)
            If Code <> 0 Then
                Code = CheckControlRecord(xFunID.Text, wType, 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If CheckControlRecord(xFunID.Text, wType, 2) = 0 Then
            SetTaskProgress("ProDGS", -500)
            SetTaskButton("BDGS", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsDGS", 848, 0)
            UpdateControlRecord(xFunID.Text, xUserID, "End", 0, "End", 0)
        Else
            SetTaskStatus("StsDGS", 890, 1)
            uJavaScript.PopMsg(Me, "程式未執行完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BRBS_Click)
    '**     執行RBS Program 
    '**
    '*****************************************************************
    Protected Sub BRBS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRBS.Click
        '
        '判斷是否完成作業
        Dim wType As String = "RBS"

        UpdateControlRecord(xFunID.Text, xUserID, "Action2", 2, wType, 1)
        '
        Dim Code As Integer
        Code = CheckControlRecord(xFunID.Text, wType, 2)
        If Code <> 0 Then
            Code = CheckControlRecord(xFunID.Text, wType, 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = CheckControlRecord(xFunID.Text, wType, 2)
            If Code <> 0 Then
                Code = CheckControlRecord(xFunID.Text, wType, 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If CheckControlRecord(xFunID.Text, wType, 2) = 0 Then
            SetTaskProgress("ProRBS", -500)
            SetTaskButton("BRBS", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsRBS", 848, 0)
            UpdateControlRecord(xFunID.Text, xUserID, "End", 0, "End", 0)
        Else
            SetTaskStatus("StsRBS", 890, 1)
            uJavaScript.PopMsg(Me, "程式未執行完成，請確認!")
        End If
        '
    End Sub
End Class
