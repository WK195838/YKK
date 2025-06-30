Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class POPMonthList
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim wUserID As String           'UserID
    '
    Dim SumWaitTime As Double = 0   '合計WaitTime 
    Dim SumProdTime As Double = 0   '合計ProdTime 
    Dim SPDNo As String = ""
    Dim ORNo As String = ""
    Dim BeforeTime As String = ""
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
        If Not Me.IsPostBack Then
            SetSearchItem()
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "POPMonthList.aspx"      '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
        wUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選條件
    '**
    '*****************************************************************
    Sub SetSearchItem()
        Dim i As Integer
        Dim ListItem1 As New ListItem
        '
        '年 
        For i = 2007 To 2023
            DStartY.Items.Add(CStr(i))
        Next
        ListItem1.Text = CStr(DateTime.Now.Year)
        ListItem1.Value = CStr(DateTime.Now.Year)
        DStartY.SelectedIndex = DStartY.Items.IndexOf(ListItem1)
        '
        '月 
        For i = 1 To 12
            DStartM.Items.Add(CStr(i))
        Next
        ListItem1.Text = CStr(DateTime.Now.Month)
        ListItem1.Value = CStr(DateTime.Now.Month)
        DStartM.SelectedIndex = DStartM.Items.IndexOf(ListItem1)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     顯示資料
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        Dim SQL As String
        '
        Dim xStartTime As String = DStartY.Text + "/" + DStartM.Text + "/" + "01"
        Dim xEndTime As String = ""
        If CInt(DStartM.Text) = 12 Then
            xEndTime = CStr(CInt(DStartY.Text) + 1) + "/" + "01" + "/" + "01"
        Else
            xEndTime = DStartY.Text + "/" + CStr(CInt(DStartM.Text) + 1) + "/" + "01"
        End If
        '
        'Delete W_POPMonthList
        SQL = "Delete From W_POPMonthList "
        uDataBase.ExecuteNonQuery(SQL)
        '
        'MsgBox("xStartTime=[" + xStartTime + "]" + Chr(13) + "end=[" + xEndTime + "]")
        '
        SQL = "Select "
        SQL &= "'SPD' as Sys, No, FormNo, FormSno, Step, Seqno, ApplyID, ApplyName "
        SQL &= "From V_WaitHandle_01 "
        SQL &= "Where FormNo  = '" & "000003" & "' "
        SQL &= "  And Step    = '" & CStr("80") & "' "
        SQL &= "  And ApplyTime >= '" & xStartTime & "' "
        SQL &= "  And ApplyTime <  '" & xEndTime & "' "
        If DNo.Text <> "" And DNo.Text <> "No" Then
            SQL &= "  And No Like '%" & DNo.Text & "%' "
        End If
        SQL &= "Group by No, FormNo, FormSno, Step, Seqno, ApplyID, ApplyName "
        SQL &= "Order by No, FormNo, FormSno, Step, Seqno, ApplyID, ApplyName "
        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
        For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
            '** ReceiptTime
            SQL = "Select ReceiptTime From T_WaitHandle "
            SQL &= "Where FormNo  = '" & dt_WaitHandle.Rows(i)("FormNo").ToString.Trim & "' "
            SQL &= "  And FormSno = '" & dt_WaitHandle.Rows(i)("FormSno").ToString & "' "
            SQL &= "  And Step    = '" & dt_WaitHandle.Rows(i)("Step").ToString & "' "
            SQL &= "  And SeqNo   = '" & dt_WaitHandle.Rows(i)("SeqNo").ToString & "' "
            SQL &= "Order by ReceiptTime "
            Dim dt_ReceiptTime As DataTable = uDataBase.GetDataTable(SQL)
            Dim wReceiptTime As String = CDate(dt_ReceiptTime.Rows(0)("ReceiptTime")).ToString("yyyy/MM/dd HH:mm") + ":00"
            '
            '** POP Information
            Dim str As String
            Dim wNo, wORNo, wPRNo, wOP, wCode, wItemName, wWaitTime, wStartTime, wEndTime, wProdTime, wMachineNo, wWaveTime As String
            Dim xStart, xEnd, xMachine As String

            Dim cn As New OleDbConnection
            Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
            Dim ds As New DataSet
            Dim dt As New DataTable

            SQL = "SELECT "
            SQL &= "LN1CX2, LN2CX2, ITMCX2, PSHNX2, ORDNX2 "
            SQL &= "FROM WAVEALIB.TWF382SP "
            SQL &= "WHERE CORNX2 = '" + dt_WaitHandle.Rows(i)("Sys").ToString.Trim + "' "
            SQL &= "  AND ORFNX2 = '" + dt_WaitHandle.Rows(i)("No").ToString.Trim + "' "
            SQL &= "ORDER BY ORDNX2, PPDUX2, PSHNX2 "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
            DBAdapter1.Fill(ds, "POP")
            dt = ds.Tables("POP")
            For j As Integer = 0 To dt.Rows.Count - 1
                '
                wNo = ""
                wORNo = ""
                wPRNo = ""
                wOP = ""
                wCode = ""
                wItemName = ""
                wWaitTime = ""
                wStartTime = ""
                wEndTime = ""
                wProdTime = ""
                wMachineNo = ""
                wWaveTime = ""
                '
                '** No
                wNo = dt_WaitHandle.Rows(i)("No").ToString.Trim
                '
                '** ORNo
                wORNo = dt.Rows(j).Item("ORDNX2").ToString.Trim
                '
                '** PRNo
                wPRNo = dt.Rows(j).Item("PSHNX2").ToString.Trim
                '
                Dim ds1 As New DataSet
                Dim dt1 As New DataTable
                '工程
                ds1.Clear()
                str = dt.Rows(j).Item("LN1CX2").ToString.Trim + "-" + dt.Rows(j).Item("LN2CX2").ToString.Trim
                SQL = "SELECT CN1I09 FROM WAVEDLIB.C0900 "
                SQL &= "WHERE DGRC09 = 'WSHC' AND DDTC09 = '" & dt.Rows(j).Item("LN1CX2").ToString.Trim & "' "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                DBAdapter2.Fill(ds1, "C0900")
                dt1 = ds1.Tables("C0900")
                If dt1.Rows.Count > 0 Then
                    str = str + "(" + dt1.Rows(0).Item("CN1I09").ToString.Trim + ")"
                End If
                wOP = str
                'Code
                wCode = dt.Rows(j).Item("ITMCX2").ToString.Trim
                'ItemName
                ds1.Clear()
                str = ""
                SQL = "SELECT IT1IA0 FROM WAVEDLIB.FA000 "
                SQL &= "WHERE ITMCA0 = '" & dt.Rows(j).Item("ITMCX2").ToString.Trim & "' "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                DBAdapter3.Fill(ds1, "FA000")
                dt1 = ds1.Tables("FA000")
                If dt1.Rows.Count > 0 Then
                    str = dt1.Rows(0).Item("IT1IA0").ToString.Trim
                End If
                wItemName = str
                '生產等待時間
                wWaitTime = ""
                '生產開始 / 生產完成 / 機台NO.
                ds1.Clear()
                str = ""
                xStart = ""
                xEnd = ""
                xMachine = ""

                SQL = "SELECT PSDUAB, SSTTAB, PCPUAB, PCMTAB, PPLCAB, PPMCAB, PPTNAB FROM WAVEDLIB.FAB00 "
                SQL &= "WHERE PSHNAB = '" & dt.Rows(j).Item("PSHNX2").ToString.Trim & "' "
                Dim DBAdapter4 As New OleDbDataAdapter(SQL, cn)
                DBAdapter4.Fill(ds1, "FAB00")
                dt1 = ds1.Tables("FAB00")
                For k As Integer = 0 To dt1.Rows.Count - 1
                    '生產開始
                    If dt1.Rows(k).Item("PSDUAB").ToString.Trim <> "0" Then
                        str = Mid(dt1.Rows(k).Item("PSDUAB").ToString.Trim, 1, 4) + "/" + _
                              Mid(dt1.Rows(k).Item("PSDUAB").ToString.Trim, 5, 2) + "/" + _
                              Mid(dt1.Rows(k).Item("PSDUAB").ToString.Trim, 7, 2)
                        If Len(dt1.Rows(k).Item("SSTTAB").ToString.Trim) > 5 Then
                            str = str + " " + _
                                  Mid(dt1.Rows(k).Item("SSTTAB").ToString.Trim, 1, 2) + ":" + _
                                  Mid(dt1.Rows(k).Item("SSTTAB").ToString.Trim, 3, 2) + ":" + _
                                  "00"
                        Else
                            If Len(dt1.Rows(k).Item("SSTTAB").ToString.Trim) > 4 Then
                                str = str + " 0" + _
                                      Mid(dt1.Rows(k).Item("SSTTAB").ToString.Trim, 1, 1) + ":" + _
                                      Mid(dt1.Rows(k).Item("SSTTAB").ToString.Trim, 2, 2) + ":" + _
                                      "00"
                            Else
                                str = str + " 00:" + _
                                      Mid(dt1.Rows(k).Item("SSTTAB").ToString.Trim, 1, 2) + ":" + _
                                      "00"
                            End If
                        End If
                        wStartTime = str
                    End If
                    '生產完成
                    If dt1.Rows(k).Item("PCPUAB").ToString.Trim <> "0" Then
                        str = Mid(dt1.Rows(k).Item("PCPUAB").ToString.Trim, 1, 4) + "/" + _
                              Mid(dt1.Rows(k).Item("PCPUAB").ToString.Trim, 5, 2) + "/" + _
                              Mid(dt1.Rows(k).Item("PCPUAB").ToString.Trim, 7, 2)
                        If Len(dt1.Rows(k).Item("PCMTAB").ToString.Trim) > 5 Then
                            str = str + " " + _
                                  Mid(dt1.Rows(k).Item("PCMTAB").ToString.Trim, 1, 2) + ":" + _
                                  Mid(dt1.Rows(k).Item("PCMTAB").ToString.Trim, 3, 2) + ":" + _
                                  "00"
                        Else
                            If Len(dt1.Rows(k).Item("PCMTAB").ToString.Trim) > 4 Then
                                str = str + " 0" + _
                                      Mid(dt1.Rows(k).Item("PCMTAB").ToString.Trim, 1, 1) + ":" + _
                                      Mid(dt1.Rows(k).Item("PCMTAB").ToString.Trim, 2, 2) + ":" + _
                                      "00"
                            Else
                                str = str + " 00:" + _
                                      Mid(dt1.Rows(k).Item("PCMTAB").ToString.Trim, 1, 2) + ":" + _
                                      "00"
                            End If
                        End If
                        wEndTime = str
                    End If
                    '機台NO.
                    wMachineNo = dt1.Rows(k).Item("PPLCAB").ToString.Trim + dt1.Rows(k).Item("PPMCAB").ToString.Trim + "(" + dt1.Rows(k).Item("PPTNAB").ToString.Trim + ")"
                Next

                '生產時間
                wProdTime = ""
                'Wave's完成日
                ds1.Clear()
                str = ""
                SQL = "SELECT PCPU9H FROM WAVEDLIB.F9H00 "
                SQL &= "WHERE PSHN9H = '" & dt.Rows(j).Item("PSHNX2").ToString.Trim & "' "
                Dim DBAdapter5 As New OleDbDataAdapter(SQL, cn)
                DBAdapter5.Fill(ds1, "F9H00")
                dt1 = ds1.Tables("F9H00")
                If dt1.Rows.Count > 0 Then
                    str = Mid(dt1.Rows(0).Item("PCPU9H").ToString.Trim, 1, 4) + "/" + _
                          Mid(dt1.Rows(0).Item("PCPU9H").ToString.Trim, 5, 2) + "/" + _
                          Mid(dt1.Rows(0).Item("PCPU9H").ToString.Trim, 7, 2)
                End If
                wWaveTime = str

                'Insert W_POPMonthList
                SQL = "INSERT INTO W_POPMonthList ( " & _
                      "No, ORNo, PRNo, OP, Code, " & _
                      "ItemName, " & _
                      "WaitTime, StartTime, EndTime, ProdTime, MachineNo, WaveTime, " & _
                      "FormNo, FormSno, Step, SeqNo, ApplyID, ApplyName, ReceiptTime, " & _
                      "CreateUser, CreateTime ) " & _
                "VALUES("
                SQL &= "'" & wNo & "', "
                SQL &= "'" & wORNo & "', "
                SQL &= "'" & wPRNo & "', "
                SQL &= "'" & wOP & "', "
                SQL &= "'" & wCode & "', "
                SQL &= "'" & wItemName & "', "
                SQL &= "'" & wWaitTime & "', "
                SQL &= "'" & wStartTime & "', "
                SQL &= "'" & wEndTime & "', "
                SQL &= "'" & wProdTime & "', "
                SQL &= "'" & wMachineNo & "', "
                SQL &= "'" & wWaveTime & "', "
                SQL &= "'" & dt_WaitHandle.Rows(i)("FormNo").ToString.Trim & "', "
                SQL &= "'" & dt_WaitHandle.Rows(i)("FormSno").ToString & "', "
                SQL &= "'" & dt_WaitHandle.Rows(i)("Step").ToString & "', "
                SQL &= "'" & dt_WaitHandle.Rows(i)("SeqNo").ToString & "', "
                SQL &= "'" & dt_WaitHandle.Rows(i)("ApplyID").ToString.Trim & "', "
                SQL &= "'" & dt_WaitHandle.Rows(i)("ApplyName").ToString.Trim & "', "
                SQL &= "'" & wReceiptTime & "', "
                SQL &= "'" & wUserID & "', "
                SQL &= "'" & NowDateTime & "') "
                uDataBase.ExecuteNonQuery(SQL)

            Next
        Next
        '
        'Show GridView
        SQL = "Select "
        SQL &= "*, "
        SQL &= "'@' As History, "

        SQL &= "'http://10.245.1.10/WorkFlow/ManufInSheet_02.aspx?' + "
        SQL &= "'pFormNo='   + FormNo + "
        SQL &= "'&pFormSno=' + str(FormSno,Len(FormSno)) "
        SQL &= " As ViewURL, "

        SQL &= "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL &= "'pFormNo='   + FormNo + "
        SQL &= "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL &= "'&pStep='    + str(Step,Len(Step)) + "
        SQL &= "'&pSeqNo='   + str(SeqNo,Len(SeqNo)) + "
        SQL &= "'&pApplyID=' + ApplyID "
        SQL &= " As OPURL "

        SQL &= "From W_POPMonthList "
        SQL &= "Order by No, ORNo, ReceiptTime, StartTime "
        Dim dt_POPMonthList As DataTable = uDataBase.GetDataTable(SQL)
            '
        GridView1.DataSource = dt_POPMonthList
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '
        'Header-------------------------------------------------------
        If e.Row.RowType = DataControlRowType.Header Then
            SPDNo = ""
            ORNo = ""
            SumWaitTime = 0
            SumProdTime = 0
        End If
        '
        'Data-------------------------------------------------------
        If e.Row.RowType = DataControlRowType.DataRow Then
            '
            If e.Row.Cells(9).Text.ToString <> "&nbsp;" And e.Row.Cells(10).Text.ToString <> "&nbsp;" Then
                Dim str As String = ""
                Dim xTime As Integer
                '
                If e.Row.Cells(0).Text.ToString <> SPDNo Or e.Row.Cells(2).Text.ToString <> ORNo Then
                    BeforeTime = e.Row.Cells(7).Text.ToString

                    SPDNo = e.Row.Cells(0).Text.ToString
                    ORNo = e.Row.Cells(2).Text.ToString
                Else
                    '
                    '生產等待時間
                    'MsgBox("Before=[" + BeforeTime + "]" + Chr(13) + "等待=[" + e.Row.Cells(9).Text.ToString + "]")
                    '
                    str = ""
                    xTime = DateDiff("n", CDate(BeforeTime), CDate(e.Row.Cells(9).Text.ToString))
                    'If xTime >= 60 Then
                    '    str = CStr(Fix(xTime / 60)) + ":" + CStr(xTime - Fix(xTime / 60) * 60)
                    'Else
                    '    str = "0:" + CStr(xTime)
                    'End If
                    e.Row.Cells(8).Text = CStr(xTime)
                    SumWaitTime = SumWaitTime + xTime
                End If
                '
                '生產時間
                'MsgBox("Start=[" + e.Row.Cells(4).Text.ToString + "]" + Chr(13) + "End=[" + e.Row.Cells(5).Text.ToString + "]")
                '
                str = ""
                xTime = DateDiff("n", CDate(e.Row.Cells(9).Text.ToString), CDate(e.Row.Cells(10).Text.ToString))
                'If xTime >= 60 Then
                '    str = CStr(Fix(xTime / 60)) + ":" + CStr(xTime - Fix(xTime / 60) * 60)
                'Else
                '    str = "0:" + CStr(xTime)
                'End If
                e.Row.Cells(11).Text = CStr(xTime)
                SumProdTime = SumProdTime + xTime
                '
                'Keep BeforeTime
                BeforeTime = e.Row.Cells(10).Text.ToString
            End If

        End If
        '
        'Footer-------------------------------------------------------
        If e.Row.RowType = DataControlRowType.Footer Then
            'Dim str As String = ""
            ''
            'e.Row.Cells(2).Text = "合計"
            ''生產等待
            'If SumWaitTime >= 60 Then
            '    str = CStr(Fix(SumWaitTime / 60)) + ":" + CStr(SumWaitTime - Fix(SumWaitTime / 60) * 60)
            'Else
            '    str = "0:" + CStr(SumWaitTime)
            'End If
            'e.Row.Cells(3).Text = str
            ''生產
            'If SumProdTime >= 60 Then
            '    str = CStr(Fix(SumProdTime / 60)) + ":" + CStr(SumProdTime - Fix(SumProdTime / 60) * 60)
            'Else
            '    str = "0:" + CStr(SumProdTime)
            'End If
            'e.Row.Cells(6).Text = str
        End If
    End Sub


End Class
