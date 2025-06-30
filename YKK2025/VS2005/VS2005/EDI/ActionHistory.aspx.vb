Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing

Partial Class ActionHistory
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '
    Dim NowDateTime As String       '現在日期時間
    Dim xUserID, xLogID, xBuyer, xAction
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
        '
        xLogID = Request.QueryString("pLogID")
        xBuyer = Request.QueryString("pBuyer")
        xUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetSearchItem)
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchItem()
        ' Customer & Buyer & LogID
        DCustBuyer.Text = xBuyer
        DLogID.Text = xLogID
        If UCase(xUserID) <> "IT003" Then
            DCustBuyer.ReadOnly = True
            DLogID.ReadOnly = True
        End If
        ' 狀態
        DSts.Items.Clear()
        Dim ListItem1, ListItem2, ListItem3, ListItem4 As New ListItem
        'ALL
        ListItem1.Text = "ALL"
        ListItem1.Value = "ALL"
        ListItem1.Selected = False
        DSts.Items.Add(ListItem1)
        '正常
        ListItem2.Text = "正常"
        ListItem2.Value = "0"
        ListItem2.Selected = False
        DSts.Items.Add(ListItem2)
        '異常
        ListItem3.Text = "異常"
        ListItem3.Value = "1"
        ListItem3.Selected = True
        DSts.Items.Add(ListItem3)
        '
        ' Function
        DAction.Items.Clear()
        'ALL
        ListItem4.Text = "ALL"
        ListItem4.Value = "ALL"
        ListItem4.Selected = True
        DAction.Items.Add(ListItem4)
        '
        Dim i As Integer
        Dim Sql As String
        Sql = "SELECT "
        Sql = Sql + "Action, "
        Sql = Sql + "Case Action "
        Sql = Sql + "When 'Rule2Data' Then '匯入' "
        Sql = Sql + "When 'MakePONO' Then '製作採購號碼' "
        Sql = Sql + "When 'CheckPONO' Then '檢測採購號碼' "
        Sql = Sql + "When 'MakeGRPC' Then '製作Group Code' "
        Sql = Sql + "When 'CheckGRPC' Then '檢測Group Code' "
        Sql = Sql + "When 'CheckCompanyCode' Then '檢測Company Code' "
        Sql = Sql + "When 'CheckKeepCode' Then '檢測Keep Code' "
        Sql = Sql + "When 'MakeColorCode' Then '製作Color Code' "
        Sql = Sql + "When 'CheckColorCode' Then '檢測Color Code' "
        Sql = Sql + "When 'CheckItemCode' Then '檢測Item Code' "
        Sql = Sql + "When 'CheckDuplicateData' Then '檢測重覆資料' "
        Sql = Sql + "When 'CheckNikeVDP' Then '檢測NIKE-VDP' "
        Sql = Sql + "When 'CheckPOStructure' Then '檢測PO結構資料' "
        Sql = Sql + "Else '轉Waves' End As ActionDesc "
        Sql = Sql + "From L_ActionHistory "
        Sql = Sql + "Where LogID = '" + xLogID + "' "
        Sql = Sql + "  And Buyer = '" + xBuyer + "' "
        Sql = Sql + "Group by Action "
        Sql = Sql + "Order by Action "
        Dim dt_Action As DataTable = uDataBase.GetDataTable(Sql)
        For i = 0 To dt_Action.Rows.Count - 1
            Dim ListItem5 As New ListItem
            ListItem5.Text = dt_Action.Rows(i).Item("ActionDesc")
            ListItem5.Value = dt_Action.Rows(i).Item("Action")
            ListItem5.Selected = False
            DAction.Items.Add(ListItem5)
        Next
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     資料抽出
    '**
    '*****************************************************************
    Sub DataList()
        Dim Sql, xPO, xSeq, xType As String
        '
        Sql = "Select BuyerGroup From M_ControlRecord "
        Sql &= "Where Buyer = '" & DCustBuyer.Text & "' "
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(Sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            If fpObj.GetFunctionCode(dt_ControlRecord.Rows(0).Item("BuyerGroup").ToString, 2) = "P" Then        ' GRBuyer(2)=P (使用自動PO)
                xType = "P"
            Else
                xType = "X"
            End If
        End If
        '
        Dim OleDbConnection1 As New OleDbConnection
        Dim ds1 As New DataSet
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '
        Sql = "SELECT "
        Sql = Sql + "'' As URL, '' As URL1, "
        Sql = Sql + "'.....' As Err1, '.....' As Err2, "
        Sql = Sql + "LogID, Buyer, "
        '
        Sql = Sql + "Case Action "
        Sql = Sql + "When 'Rule2Data' Then '匯入' "
        Sql = Sql + "When 'MakePONO' Then '製作採購號碼' "
        Sql = Sql + "When 'CheckPONO' Then '檢測採購號碼' "
        Sql = Sql + "When 'MakeGRPC' Then '製作Group Code' "
        Sql = Sql + "When 'CheckGRPC' Then '檢測Group Code' "
        Sql = Sql + "When 'CheckCompanyCode' Then '檢測Company Code' "
        Sql = Sql + "When 'CheckKeepCode' Then '檢測Keep Code' "
        Sql = Sql + "When 'MakeColorCode' Then '製作Color Code' "
        Sql = Sql + "When 'CheckColorCode' Then '檢測Color Code' "
        Sql = Sql + "When 'CheckItemCode' Then '檢測Item Code' "
        Sql = Sql + "When 'CheckDuplicateData' Then '檢測重覆資料' "
        Sql = Sql + "When 'CheckNikeVDP' Then '檢測NIKE-VDP' "
        Sql = Sql + "When 'CheckPOStructure' Then '檢測PO結構資料' "
        Sql = Sql + "When 'GetSalesPrice' Then '取得單價Inf.' "
        Sql = Sql + "Else '轉Waves' End As ActionDesc, "
        '
        Sql = Sql + "Convert(VARCHAR(20), RunTime, 120) As RunTimeDesc, "
        Sql = Sql + "Case Error When 0 Then '正常' Else '異常' End As ErrorDesc, "
        Sql = Sql + "D1, D2, D3, D4, D5 "
        Sql = Sql + "From L_ActionHistory "
        Sql = Sql + "Where LogID = '" + DLogID.Text + "' "
        Sql = Sql + "  And Buyer = '" + DCustBuyer.Text + "' "
        '狀態
        If DSts.SelectedValue <> "ALL" Then
            Sql = Sql + "And Error =  '" + DSts.SelectedValue + "' "
        End If
        'Action
        If DAction.SelectedValue <> "ALL" Then
            Sql = Sql + "And Action =  '" + DAction.SelectedValue + "' "
        End If
        'Line
        If DLine.Text <> "" Then
            Sql = Sql + "And D1 = '" + DLine.Text + "' "
        End If
        Sql = Sql + "Order by Unique_ID "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, OleDbConnection1)
        DBAdapter1.Fill(ds1, "ActionLog")
        If ds1.Tables("ActionLog").Rows.Count > 0 Then
            For i As Integer = 0 To ds1.Tables("ActionLog").Rows.Count - 1
                xPO = ""
                xSeq = ""
                '
                If xType = "P" Then        ' GRBuyer(2)=P (使用自動PO) 
                    Sql = "SELECT D4 "
                    Sql &= "From L_ActionHistory "
                    Sql &= "Where LogID = '" + DLogID.Text + "' "
                    Sql &= "  And Buyer = '" + DCustBuyer.Text + "' "
                    Sql &= "  And D1 = '" + ds1.Tables("ActionLog").Rows(i)("D1") + "' "
                    Sql &= "  And D2 = 'CORN5K' "
                    Dim dt_Log1 As DataTable = uDataBase.GetDataTable(Sql)
                    If dt_Log1.Rows.Count > 0 Then
                        xPO = dt_Log1.Rows(0).Item("D4")
                    End If
                Else
                    Sql = "SELECT D4 "
                    Sql &= "From L_ActionHistory "
                    Sql &= "Where LogID = '" + DLogID.Text + "' "
                    Sql &= "  And Buyer = '" + DCustBuyer.Text + "' "
                    Sql &= "  And D1 = '" + ds1.Tables("ActionLog").Rows(i)("D1") + "' "
                    Sql &= "  And D2 = 'PODN5K' "
                    Dim dt_Log1 As DataTable = uDataBase.GetDataTable(Sql)
                    If dt_Log1.Rows.Count > 0 Then
                        xPO = dt_Log1.Rows(0).Item("D4")
                    End If
                End If
                '
                Sql = "SELECT D4 "
                Sql &= "From L_ActionHistory "
                Sql &= "Where LogID = '" + DLogID.Text + "' "
                Sql &= "  And Buyer = '" + DCustBuyer.Text + "' "
                Sql &= "  And D1 = '" + ds1.Tables("ActionLog").Rows(i)("D1") + "' "
                Sql &= "  And D2 = 'GRPC5K' "
                Dim dt_Log2 As DataTable = uDataBase.GetDataTable(Sql)
                If dt_Log2.Rows.Count > 0 Then
                    xSeq = dt_Log2.Rows(0).Item("D4")
                End If
                '
                If xPO <> "" And xSeq <> "" Then
                    ds1.Tables("ActionLog").Rows(i)(0) = "InfCustPOList.aspx?" + "pUserID=" + xUserID + "&pBuyer=" + DCustBuyer.Text + "&pType=" + xType + "&pPO=" + xPO + "&pSeq=" + xSeq
                End If
                '
                ds1.Tables("ActionLog").Rows(i)(1) = "CheckWavesDataList.aspx?pBuyer=" + DCustBuyer.Text + "&pLogID=" + DLogID.Text
            Next
            '
            GridView1.DataSource = ds1
            GridView1.DataBind()
        End If
        '
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     按鈕(Go)
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataList()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     顏色處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim xBackColor As Boolean = False
            '--異常--------------------------------------------
            If e.Row.Cells(4).Text.ToString = "異常" Then
                e.Row.ForeColor = Color.Blue
                e.Row.BackColor = Color.LightPink
            End If
        End If
    End Sub


End Class
