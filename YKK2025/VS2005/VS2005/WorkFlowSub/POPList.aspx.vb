Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class POPList
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim k_SYS As String             'Key --> 系統(SPD/SCD)
    Dim k_NO As String              'Key --> No(委託No)
    Dim k_FormNo As String
    Dim k_Formsno As Integer
    Dim k_Step As Integer
    '
    Dim SumWaitTime As Double = 0   '合計WaitTime 
    Dim SumProdTime As Double = 0   '合計ProdTime 
    Dim xBeforeTime As String = ""
    Dim FirstData As Boolean = False
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
        SetPopupFunction()                          '設定彈出視窗事件
        If Not IsPostBack Then                      'PostBack
            ShowData()
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
        Response.Cookies("PGM").Value = "POPList.aspx"      '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日期時間
        k_SYS = Request.QueryString("pSYS")     '系統
        k_NO = Request.QueryString("pNO")       'No
        k_FormNo = Request.QueryString("pFormNo")
        k_Formsno = Request.QueryString("pFormSno")
        k_Step = Request.QueryString("pStep")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim xStart, xEnd, xMachine As String
        Dim cn As New OleDbConnection
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        Dim ds As New DataSet
        Dim dt As New DataTable
        '
        DMessage.Visible = False
        '
        SQL = "SELECT "
        SQL &= "'' As OP, "
        SQL &= "'' As Code, "
        SQL &= "'' As ItemName, "
        SQL &= "'' As StartTime, "
        SQL &= "'' As EndTime, "
        SQL &= "'' As MachineNo, "
        SQL &= "'' As WaveTime, "
        SQL &= "'' As PRNO, "
        'ADD-Start Joy 2012/8/8
        SQL &= "'' As WaitTime, "
        SQL &= "'' As ProdTime, "
        'ADD-End
        SQL &= "LN1CX2, LN2CX2, ITMCX2, PSHNX2, ORDNX2 "
        SQL &= "FROM WAVEALIB.TWF382SP "
        SQL &= "WHERE CORNX2 = '" + k_SYS + "' "
        SQL &= "  AND ORFNX2 = '" + k_NO + "' "
        SQL &= "ORDER BY ORDNX2, PPDUX2, PSHNX2 "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "POP")
        dt = ds.Tables("POP")
        If dt.Rows.Count > 0 Then
            '
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim str As String
                Dim ds1 As New DataSet
                Dim dt1 As New DataTable
                '
                '工程
                ds1.Clear()
                str = dt.Rows(i).Item("LN1CX2") + "-" + dt.Rows(i).Item("LN2CX2")
                SQL = "SELECT CN1I09 FROM WAVEDLIB.C0900 "
                SQL &= "WHERE DGRC09 = 'WSHC' AND DDTC09 = '" & dt.Rows(i).Item("LN1CX2") & "' "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                DBAdapter2.Fill(ds1, "C0900")
                dt1 = ds1.Tables("C0900")
                If dt1.Rows.Count > 0 Then
                    str = str + "(" + dt1.Rows(0).Item("CN1I09") + ")"
                End If
                dt.Rows(i)(0) = str
                'Code
                dt.Rows(i)(1) = dt.Rows(i).Item("ITMCX2")
                'ItemName
                ds1.Clear()
                str = ""
                SQL = "SELECT IT1IA0 FROM WAVEDLIB.FA000 "
                SQL &= "WHERE ITMCA0 = '" & dt.Rows(i).Item("ITMCX2") & "' "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                DBAdapter3.Fill(ds1, "FA000")
                dt1 = ds1.Tables("FA000")
                If dt1.Rows.Count > 0 Then
                    str = dt1.Rows(0).Item("IT1IA0")
                End If
                dt.Rows(i)(2) = str

                '生產開始 / 生產完成 / 機台NO.
                ds1.Clear()
                str = ""
                xStart = ""
                xEnd = ""
                xMachine = ""
                SQL = "SELECT PSDUAB, SSTTAB, PCPUAB, PCMTAB, PPLCAB, PPMCAB, PPTNAB FROM WAVEDLIB.FAB00 "
                SQL &= "WHERE PSHNAB = '" & dt.Rows(i).Item("PSHNX2") & "' "
                Dim DBAdapter4 As New OleDbDataAdapter(SQL, cn)
                DBAdapter4.Fill(ds1, "FAB00")
                dt1 = ds1.Tables("FAB00")
                For j As Integer = 0 To dt1.Rows.Count - 1
                    '生產開始
                    If dt1.Rows(j).Item("PSDUAB").ToString <> "0" Then
                        str = Mid(dt1.Rows(j).Item("PSDUAB").ToString, 1, 4) + "/" + _
                              Mid(dt1.Rows(j).Item("PSDUAB").ToString, 5, 2) + "/" + _
                              Mid(dt1.Rows(j).Item("PSDUAB").ToString, 7, 2)
                        If Len(dt1.Rows(j).Item("SSTTAB").ToString) > 5 Then
                            str = str + " " + _
                                  Mid(dt1.Rows(j).Item("SSTTAB").ToString, 1, 2) + ":" + _
                                  Mid(dt1.Rows(j).Item("SSTTAB").ToString, 3, 2) + ":" + _
                                  "00"
                        Else
                            If Len(dt1.Rows(j).Item("SSTTAB").ToString.Trim) > 4 Then
                                str = str + " 0" + _
                                      Mid(dt1.Rows(j).Item("SSTTAB").ToString.Trim, 1, 1) + ":" + _
                                      Mid(dt1.Rows(j).Item("SSTTAB").ToString.Trim, 2, 2) + ":" + _
                                      "00"
                            Else
                                str = str + " 00:" + _
                                      Mid(dt1.Rows(j).Item("SSTTAB").ToString.Trim, 1, 2) + ":" + _
                                      "00"
                            End If
                        End If
                        dt.Rows(i)(3) = str
                    End If
                    '生產完成
                    If dt1.Rows(j).Item("PCPUAB").ToString <> "0" Then
                        str = Mid(dt1.Rows(j).Item("PCPUAB").ToString, 1, 4) + "/" + _
                              Mid(dt1.Rows(j).Item("PCPUAB").ToString, 5, 2) + "/" + _
                              Mid(dt1.Rows(j).Item("PCPUAB").ToString, 7, 2)
                        If Len(dt1.Rows(j).Item("PCMTAB").ToString) > 5 Then
                            str = str + " " + _
                                  Mid(dt1.Rows(j).Item("PCMTAB").ToString, 1, 2) + ":" + _
                                  Mid(dt1.Rows(j).Item("PCMTAB").ToString, 3, 2) + ":" + _
                                  "00"
                        Else
                            If Len(dt1.Rows(j).Item("PCMTAB").ToString.Trim) > 4 Then
                                str = str + " 0" + _
                                      Mid(dt1.Rows(j).Item("PCMTAB").ToString.Trim, 1, 1) + ":" + _
                                      Mid(dt1.Rows(j).Item("PCMTAB").ToString.Trim, 2, 2) + ":" + _
                                      "00"
                            Else
                                str = str + " 00:" + _
                                      Mid(dt1.Rows(j).Item("PCMTAB").ToString.Trim, 1, 2) + ":" + _
                                      "00"
                            End If
                        End If
                        dt.Rows(i)(4) = str
                    End If
                    '機台NO.
                    str = dt1.Rows(j).Item("PPLCAB") + dt1.Rows(j).Item("PPMCAB") + "(" + dt1.Rows(j).Item("PPTNAB") + ")"
                    dt.Rows(i)(5) = str
                Next
                'Wave's完成日
                ds1.Clear()
                str = ""
                SQL = "SELECT PCPU9H FROM WAVEDLIB.F9H00 "
                SQL &= "WHERE PSHN9H = '" & dt.Rows(i).Item("PSHNX2") & "' "
                Dim DBAdapter5 As New OleDbDataAdapter(SQL, cn)
                DBAdapter5.Fill(ds1, "F9H00")
                dt1 = ds1.Tables("F9H00")
                If dt1.Rows.Count > 0 Then
                    str = Mid(dt1.Rows(0).Item("PCPU9H").ToString, 1, 4) + "/" + _
                          Mid(dt1.Rows(0).Item("PCPU9H").ToString, 5, 2) + "/" + _
                          Mid(dt1.Rows(0).Item("PCPU9H").ToString, 7, 2)
                End If
                dt.Rows(i)(6) = str
                'PR-NO
                dt.Rows(i)(7) = dt.Rows(i).Item("PSHNX2") + " / " + dt.Rows(i).Item("ORDNX2") + " / " + k_NO
                '生產等待時間
                dt.Rows(i)(8) = ""
                '生產時間
                dt.Rows(i)(9) = ""
            Next
            '
            Dim dv As New DataView
            dv = dt.DefaultView
            dv.Sort = "ORDNX2, StartTime "
            '
            GridView1.DataSource = dv
            GridView1.DataBind()
        Else
            GridView1.Visible = False
            DMessage.Visible = True
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '
        '隱藏欄位(Order-No)
        e.Row.Cells(10).Visible = False
        '
        'Header-------------------------------------------------------
        If e.Row.RowType = DataControlRowType.Header Then
            Dim sql As String = ""
            Sql = "Select ReceiptTime "
            Sql = Sql & "From T_WaitHandle "
            Sql = Sql + "Where FormNo  = '" & k_FormNo & "' "
            Sql = Sql + "  And FormSno = '" & CStr(k_Formsno) & "' "
            Sql = Sql + "  And Step    = '" & CStr(k_Step) & "' "
            Sql = Sql & "Order by ReceiptTime "
            Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(Sql)
            If dt_WaitHandle.Rows.Count > 0 Then
                xBeforeTime = CDate(dt_WaitHandle.Rows(0)("ReceiptTime")).ToString("yyyy/MM/dd HH:mm") + ":00"
                FirstData = True
            End If
            '
            SumWaitTime = 0
            SumProdTime = 0
        End If
        '
        'Data-------------------------------------------------------
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim str As String = ""
            Dim xTime As Integer
            '生產等待時間
            If e.Row.Cells(4).Text.ToString <> "&nbsp;" Then
                'MsgBox("Before=[" + xBeforeTime + "]" + Chr(13) + "等待=[" + e.Row.Cells(4).Text.ToString + "]")
                '
                If Not FirstData Then
                    str = ""
                    xTime = DateDiff("n", CDate(xBeforeTime), CDate(e.Row.Cells(4).Text.ToString))
                    'If xTime >= 60 Then
                    '    str = CStr(Fix(xTime / 60)) + ":" + CStr(xTime - Fix(xTime / 60) * 60)
                    'Else
                    '    str = "0:" + CStr(xTime)
                    'End If
                    e.Row.Cells(3).Text = CStr(xTime)
                    SumWaitTime = SumWaitTime + xTime
                Else
                    e.Row.Cells(3).Text = ""
                    SumWaitTime = SumWaitTime + 0
                End If
                '
            End If
            '生產時間
            If e.Row.Cells(4).Text.ToString <> "&nbsp;" And e.Row.Cells(5).Text.ToString <> "&nbsp;" Then
                'MsgBox("Start=[" + e.Row.Cells(4).Text.ToString + "]" + Chr(13) + "End=[" + e.Row.Cells(5).Text.ToString + "]")
                '
                str = ""
                xTime = DateDiff("n", CDate(e.Row.Cells(4).Text.ToString), CDate(e.Row.Cells(5).Text.ToString))
                'If xTime >= 60 Then
                '    str = CStr(Fix(xTime / 60)) + ":" + CStr(xTime - Fix(xTime / 60) * 60)
                'Else
                '    str = "0:" + CStr(xTime)
                'End If
                e.Row.Cells(6).Text = CStr(xTime)
                SumProdTime = SumProdTime + xTime
            End If
            'Keep xBeforeTime
            If e.Row.Cells(5).Text.ToString <> "&nbsp;" Then
                xBeforeTime = e.Row.Cells(5).Text.ToString
                FirstData = False
            End If
        End If
        '
        'Footer-------------------------------------------------------
        If e.Row.RowType = DataControlRowType.Footer Then
            Dim str As String = ""
            '
            e.Row.Cells(2).Text = "合計"
            '生產等待
            'If SumWaitTime >= 60 Then
            '    str = CStr(Fix(SumWaitTime / 60)) + ":" + CStr(SumWaitTime - Fix(SumWaitTime / 60) * 60)
            'Else
            '    str = "0:" + CStr(SumWaitTime)
            'End If
            e.Row.Cells(3).Text = CStr(SumWaitTime)
            '生產
            'If SumProdTime >= 60 Then
            '    str = CStr(Fix(SumProdTime / 60)) + ":" + CStr(SumProdTime - Fix(SumProdTime / 60) * 60)
            'Else
            '    str = "0:" + CStr(SumProdTime)
            'End If
            e.Row.Cells(6).Text = CStr(SumProdTime)
        End If


    End Sub
End Class
