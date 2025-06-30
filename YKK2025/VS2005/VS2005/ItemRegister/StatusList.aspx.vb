Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class StatusList
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單No
    Dim wNo As String               'No
    Dim wKeepData As String         '封存
    Dim NowDateTime As String       '現在日期時間

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        TopPosition()                   ' 設定top
        If Not IsPostBack Then
            ShowFlowData()              ' 顯示表單資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")              '表單No
        wNo = Request.QueryString("pNo")                      'No
        wKeepData = Request.QueryString("pKeepData")  '封存
        '
        Response.Cookies("PGM").Value = "StatusList.aspx"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 1
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowFlowData)
    '**     顯示流程資料
    '**
    '*****************************************************************
    Sub ShowFlowData()
        Dim SQL As String
        Dim xZIPFormNo As String = ""
        Dim xZIPNo As String = ""
        Dim xSLDFormNo As String = ""
        Dim xSLDNo As String = ""
        Dim xCHFormNo As String = ""
        Dim xCHNo As String = ""
        If wNo <> "" Then
            '取得Key欄位
            SQL = "Select "
            SQL = SQL + "ZIPApply, ZIPFormNo, ZIPNo, "
            SQL = SQL + "SLDApply, SLDFormNo, SLDNo, "
            SQL = SQL + "CHApply, CHFormNo, CHNo "
            SQL = SQL + "From F_ItemRegisterSheet "
            SQL = SQL + "Where No= '" + wNo + "'"
            Dim dt_ItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
            If dt_ItemRegisterSheet.Rows.Count > 0 Then
                'ZIP-Key
                If dt_ItemRegisterSheet.Rows(0)("ZIPApply") = 1 Then
                    xZIPFormNo = dt_ItemRegisterSheet.Rows(0)("ZIPFormNo").ToString.Trim
                    xZIPNo = dt_ItemRegisterSheet.Rows(0)("ZIPNo").ToString.Trim
                End If
                'SLD-Key
                If dt_ItemRegisterSheet.Rows(0)("SLDApply") = 1 Then
                    xSLDFormNo = dt_ItemRegisterSheet.Rows(0)("SLDFormNo").ToString.Trim
                    xSLDNo = dt_ItemRegisterSheet.Rows(0)("SLDNo").ToString.Trim
                End If
                'CH-Key
                If dt_ItemRegisterSheet.Rows(0)("CHApply") = 1 Then
                    xCHFormNo = dt_ItemRegisterSheet.Rows(0)("CHFormNo").ToString.Trim
                    xCHNo = dt_ItemRegisterSheet.Rows(0)("CHNo").ToString.Trim
                End If
            End If
            '
            '1151-ItemRegisterSheet(營業)
            '取得流程資料
            SQL = "SELECT "
            SQL = SQL + "'流程' As Field1, "
            SQL = SQL + "StepNameDesc As Field2, "
            SQL = SQL + "DecideName As Field3, "
            SQL = SQL + "AgentName As Field4, "
            SQL = SQL + "FlowTypeDesc As Field5, "
            SQL = SQL + "DelaySts As Field6, "
            SQL = SQL + "StsDesc As Field7, "
            SQL = SQL + "DecideDescA As Field8, "
            SQL = SQL + "'預定開始：[' + BStartTimeDesc + '], ' + "
            SQL = SQL + "'預定完成：[' + BEndTimeDesc + '], ' + "
            SQL = SQL + "'實際開始：[' + AStartTimeDesc + '], ' + "
            SQL = SQL + "'實際完成：[' + AEndTimeDesc + '] ' As Field9 "
            If wKeepData = "1" Then
                SQL = SQL + "FROM V_WaitHandle_OLD_01 "
            Else
                SQL = SQL + "FROM V_WaitHandle_01B "
            End If
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And No  = '" & wNo & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc "
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
            GridView1.Style("top") = Top & "px"
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            Top = Top + (dt_WaitHandle.Rows.Count + 1) * 75
            '
            '1152-ItemRegisterSheet(ZIP)
            If xZIPNo <> "" Then
                '取得流程資料
                SQL = "SELECT "
                SQL = SQL + "'流程' As Field1, "
                SQL = SQL + "StepNameDesc As Field2, "
                SQL = SQL + "DecideName As Field3, "
                SQL = SQL + "AgentName As Field4, "
                SQL = SQL + "FlowTypeDesc As Field5, "
                SQL = SQL + "DelaySts As Field6, "
                SQL = SQL + "StsDesc As Field7, "
                SQL = SQL + "DecideDescA As Field8, "
                SQL = SQL + "'預定開始：[' + BStartTimeDesc + '], ' + "
                SQL = SQL + "'預定完成：[' + BEndTimeDesc + '], ' + "
                SQL = SQL + "'實際開始：[' + AStartTimeDesc + '], ' + "
                SQL = SQL + "'實際完成：[' + AEndTimeDesc + '] ' As Field9 "
                If wKeepData = "1" Then
                    SQL = SQL + "FROM V_WaitHandle_OLD_01 "
                Else
                    SQL = SQL + "FROM V_WaitHandle_01B "
                End If
                SQL = SQL + "Where FormNo  = '" & xZIPFormNo & "' "
                SQL = SQL + "  And No  = '" & xZIPNo & "' "
                'SQL = SQL + "Order by Unique_ID Desc "
                SQL = SQL + "Order by CreateTime Desc "
                Dim dt_WaitHandleZIP As DataTable = uDataBase.GetDataTable(SQL)
                GridView2.DataSource = uDataBase.GetDataTable(SQL)
                GridView2.DataBind()
                GridView2.Style("top") = Top & "px"
                Top = Top + (dt_WaitHandleZIP.Rows.Count + 1) * 75
            End If
            '
            '1153-ItemRegisterSheet(SLD)
            If xSLDNo <> "" Then
                '取得流程資料
                SQL = "SELECT "
                SQL = SQL + "'流程' As Field1, "
                SQL = SQL + "StepNameDesc As Field2, "
                SQL = SQL + "DecideName As Field3, "
                SQL = SQL + "AgentName As Field4, "
                SQL = SQL + "FlowTypeDesc As Field5, "
                SQL = SQL + "DelaySts As Field6, "
                SQL = SQL + "StsDesc As Field7, "
                SQL = SQL + "DecideDescA As Field8, "
                SQL = SQL + "'預定開始：[' + BStartTimeDesc + '], ' + "
                SQL = SQL + "'預定完成：[' + BEndTimeDesc + '], ' + "
                SQL = SQL + "'實際開始：[' + AStartTimeDesc + '], ' + "
                SQL = SQL + "'實際完成：[' + AEndTimeDesc + '] ' As Field9 "
                If wKeepData = "1" Then
                    SQL = SQL + "FROM V_WaitHandle_OLD_01 "
                Else
                    SQL = SQL + "FROM V_WaitHandle_01B "
                End If
                SQL = SQL + "Where FormNo  = '" & xSLDFormNo & "' "
                SQL = SQL + "  And No  = '" & xSLDNo & "' "
                'SQL = SQL + "Order by Unique_ID Desc "
                SQL = SQL + "Order by CreateTime Desc "
                Dim dt_WaitHandleSLD As DataTable = uDataBase.GetDataTable(SQL)
                GridView3.DataSource = uDataBase.GetDataTable(SQL)
                GridView3.DataBind()
                GridView3.Style("top") = Top & "px"
                Top = Top + (dt_WaitHandleSLD.Rows.Count + 1) * 75
            End If
            '
            '1154-ItemRegisterSheet(CH)
            If xCHNo <> "" Then
                '取得流程資料
                SQL = "SELECT "
                SQL = SQL + "'流程' As Field1, "
                SQL = SQL + "StepNameDesc As Field2, "
                SQL = SQL + "DecideName As Field3, "
                SQL = SQL + "AgentName As Field4, "
                SQL = SQL + "FlowTypeDesc As Field5, "
                SQL = SQL + "DelaySts As Field6, "
                SQL = SQL + "StsDesc As Field7, "
                SQL = SQL + "DecideDescA As Field8, "
                SQL = SQL + "'預定開始：[' + BStartTimeDesc + '], ' + "
                SQL = SQL + "'預定完成：[' + BEndTimeDesc + '], ' + "
                SQL = SQL + "'實際開始：[' + AStartTimeDesc + '], ' + "
                SQL = SQL + "'實際完成：[' + AEndTimeDesc + '] ' As Field9 "
                If wKeepData = "1" Then
                    SQL = SQL + "FROM V_WaitHandle_OLD_01 "
                Else
                    SQL = SQL + "FROM V_WaitHandle_01B "
                End If
                SQL = SQL + "Where FormNo  = '" & xCHFormNo & "' "
                SQL = SQL + "  And No  = '" & xCHNo & "' "
                'SQL = SQL + "Order by Unique_ID Desc "
                SQL = SQL + "Order by CreateTime Desc "
                Dim dt_WaitHandleCH As DataTable = uDataBase.GetDataTable(SQL)
                GridView4.DataSource = uDataBase.GetDataTable(SQL)
                GridView4.DataBind()
                GridView4.Style("top") = Top & "px"
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1--合併儲存格/設定背景顏色
    '**
    '*****************************************************************
    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender
        Dim i As Integer = 1
        Dim wRowSpan As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(1) As Integer
        mergeColumns(0) = 0

        '合併儲存格(目前只使用 1.日期, 2.星期)
        For MergeIdx = 0 To 0
            i = 1
            For Each mySingleRow In GridView1.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0
                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "" Then  '空白是否-("&nbsp;")
                            '合併處理
                            If mergeColumns(MergeIdx) = 0 Then
                                GridView1.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                                mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                                i = i + 1
                            End If
                        End If   '空白是否
                    Else
                        GridView1.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                        i = 1
                    End If
                Else  '資料筆數>0
                    mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                End If
            Next
        Next

        '設定背景顏色
        For Each mySingleRow In GridView1.Rows
            If mySingleRow.RowIndex = 0 Then
                mySingleRow.Cells(0).BackColor = Drawing.Color.Silver
                mySingleRow.Cells(1).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(2).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(3).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(4).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(5).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(6).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(7).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(8).BackColor = Drawing.Color.LightPink
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView2--合併儲存格/設定背景顏色
    '**
    '*****************************************************************
    Protected Sub GridView2_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.PreRender
        Dim i As Integer = 1
        Dim wRowSpan As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(1) As Integer
        mergeColumns(0) = 0

        '合併儲存格(目前只使用 1.日期, 2.星期)
        For MergeIdx = 0 To 0
            i = 1
            For Each mySingleRow In GridView2.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0
                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView2.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "" Then  '空白是否-("&nbsp;")
                            '合併處理
                            If mergeColumns(MergeIdx) = 0 Then
                                GridView2.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                                mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                                i = i + 1
                            End If
                        End If   '空白是否
                    Else
                        GridView2.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                        i = 1
                    End If
                Else  '資料筆數>0
                    mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                End If
            Next
        Next

        '設定背景顏色
        For Each mySingleRow In GridView2.Rows
            If mySingleRow.RowIndex = 0 Then
                mySingleRow.Cells(0).BackColor = Drawing.Color.Silver
                mySingleRow.Cells(1).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(2).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(3).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(4).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(5).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(6).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(7).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(8).BackColor = Drawing.Color.LightPink
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView3--合併儲存格/設定背景顏色
    '**
    '*****************************************************************
    Protected Sub GridView3_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView3.PreRender
        Dim i As Integer = 1
        Dim wRowSpan As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(1) As Integer
        mergeColumns(0) = 0

        '合併儲存格(目前只使用 1.日期, 2.星期)
        For MergeIdx = 0 To 0
            i = 1
            For Each mySingleRow In GridView3.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0
                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView3.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "" Then  '空白是否-("&nbsp;")
                            '合併處理
                            If mergeColumns(MergeIdx) = 0 Then
                                GridView3.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                                mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                                i = i + 1
                            End If
                        End If   '空白是否
                    Else
                        GridView3.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                        i = 1
                    End If
                Else  '資料筆數>0
                    mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                End If
            Next
        Next

        '設定背景顏色
        For Each mySingleRow In GridView3.Rows
            If mySingleRow.RowIndex = 0 Then
                mySingleRow.Cells(0).BackColor = Drawing.Color.Silver
                mySingleRow.Cells(1).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(2).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(3).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(4).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(5).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(6).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(7).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(8).BackColor = Drawing.Color.LightPink
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView4--合併儲存格/設定背景顏色
    '**
    '*****************************************************************
    Protected Sub GridView4_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView4.PreRender
        Dim i As Integer = 1
        Dim wRowSpan As Integer = 1
        Dim MergeIdx As Integer = 0
        Dim mySingleRow As GridViewRow
        Dim mergeColumns(1) As Integer
        mergeColumns(0) = 0

        '合併儲存格(目前只使用 1.日期, 2.星期)
        For MergeIdx = 0 To 0
            i = 1
            For Each mySingleRow In GridView4.Rows
                If CInt(mySingleRow.RowIndex) > 0 Then  '資料筆數>0
                    If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() = GridView4.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).Text.Trim() Then

                        If mySingleRow.Cells(mergeColumns(MergeIdx)).Text.Trim() <> "" Then  '空白是否-("&nbsp;")
                            '合併處理
                            If mergeColumns(MergeIdx) = 0 Then
                                GridView4.Rows(CInt(mySingleRow.RowIndex) - i).Cells(mergeColumns(MergeIdx)).RowSpan += 1
                                mySingleRow.Cells(mergeColumns(MergeIdx)).Visible = False
                                i = i + 1
                            End If
                        End If   '空白是否
                    Else
                        GridView4.Rows(CInt(mySingleRow.RowIndex)).Cells(mergeColumns(MergeIdx)).RowSpan = 1
                        i = 1
                    End If
                Else  '資料筆數>0
                    mySingleRow.Cells(mergeColumns(MergeIdx)).RowSpan = 1
                End If
            Next
        Next

        '設定背景顏色
        For Each mySingleRow In GridView4.Rows
            If mySingleRow.RowIndex = 0 Then
                mySingleRow.Cells(0).BackColor = Drawing.Color.Silver
                mySingleRow.Cells(1).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(2).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(3).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(4).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(5).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(6).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(7).BackColor = Drawing.Color.LightPink
                mySingleRow.Cells(8).BackColor = Drawing.Color.LightPink
            End If
        Next
    End Sub
End Class
