Imports System.Data
Imports System.Data.OleDb

Public Class SPDNewCommission
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DNewSheet As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSimulation As System.Web.UI.WebControls.Button
    Protected WithEvents DNew As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SPDNewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        If Not Me.IsPostBack Then
            SetNewSheet()
            SetLevel()
        End If
    End Sub

    Sub SetLevel()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        Dim wFormNo As String = DNewSheet.SelectedValue          '表單號碼
        SQL = "SELECT * FROM M_Referp  "
        SQL = SQL + "Where Cat = '007' "
        Select Case wFormNo
            Case "000001"       '圖面委託
                SQL = SQL + "and (DKey like '0%' or DKey like 'Z%') "
            Case "000002"       '圖面修改委託
                SQL = SQL + "and (DKey like '0%' or DKey like 'Z%') "
            Case "000003"       '內製委託
                SQL = SQL + "and DKey like 'In%' "
            Case "000004"       '工程連絡單
                SQL = SQL + "and DKey like 'In%' "
            Case "000005"       '業務連絡單
                SQL = SQL + "and DKey like '0%' "
            Case "000007"       '外注委託
                SQL = SQL + "and DKey like 'Out%' "
            Case "000008"       '外注工程連絡單
                SQL = SQL + "and DKey like 'Out%' "
            Case "000009"       '外注業務連絡單
                SQL = SQL + "and DKey like '0%' "
            Case "000010"       '拉頭細目單
                SQL = SQL + "and DKey like '0%' "
            Case "000011"       '進口/客供拉頭
                SQL = SQL + "and DKey like '0%' "
            Case "000012"       '型別胴體追加委託書
                SQL = SQL + "and DKey like '0%' "
            Case "000013"       '進口/客供拉頭業務連絡單
                SQL = SQL + "and DKey like '0%' "
            Case "000014"       '表面處理委託單
                SQL = SQL + "and DKey like '0%' "
            Case "000015"       '表面處理追加委託單
                SQL = SQL + "and DKey like '0%' "
            Case "800001"
                SQL = SQL + "and DKey like '0%' "
            Case "900050"       '舊版圖面委託
                SQL = SQL + "and (DKey like '0%' or DKey like 'Z%') "
            Case Else
                SQL = SQL + "and DKey like 'Other%' "
        End Select
        SQL = SQL + "Order by Cat, DKey, Data "

        OleDbConnection1.Open()
        DLevel.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DLevel.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     材質點選後事件
    '**
    '*****************************************************************
    Private Sub DNewSheet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DNewSheet.SelectedIndexChanged
        SetLevel()
    End Sub

    Sub SetNewSheet()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000' "
        SQL = SQL + "         And (IniAuthority = '0'  "
        SQL = SQL + "             Or (IniAuthority = '1' "
        SQL = SQL + "                And (   IniUser     like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    Or  IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    ) "
        SQL = SQL + "                ) "
        SQL = SQL + "             ) "
        SQL = SQL + "        ) "
        SQL = SQL + "        or "
        SQL = SQL + "        (FormNo >= '900050' And FormNo <= '900999' "
        SQL = SQL + "         And (IniAuthority = '0'  "
        SQL = SQL + "             Or (IniAuthority = '1' "
        SQL = SQL + "                And (   IniUser     like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    Or  IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                    ) "
        SQL = SQL + "                ) "
        SQL = SQL + "             ) "
        SQL = SQL + "        ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormName "

        OleDbConnection1.Open()
        DNewSheet.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FORM")
        DBTable1 = DBDataSet1.Tables("M_FORM")

        DNew.Visible = False
        DSimulation.Visible = False
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            DNewSheet.Items.Add(ListItem1)
            DNew.Visible = True
            DSimulation.Visible = True
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub DNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DNew.Click
        Dim wFormNo As String = DNewSheet.SelectedValue          '表單號碼
        Dim wFormSno As Integer = 0       '表單流水號
        Dim wStep As Integer = 1          '工程代碼
        Dim SPDStep As Integer = 1        'SPD工程代碼
        Dim wSeqNo As Integer = 0         '序號
        Dim URL As String = ""            'URL

        If Request.QueryString("pUserID") <> "" Then
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()

            SQL = "Select * From M_Users "
            SQL = SQL & " Where Active =  '1' "
            SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                If DBDataSet1.Tables("M_Users").Rows(0).Item("DivID") = "1118211" Then SPDStep = 2 '台中
                If DBDataSet1.Tables("M_Users").Rows(0).Item("DivID") = "1018950" Then SPDStep = 2 'EA
                'DB連結關閉
                OleDbConnection1.Close()

                Select Case wFormNo
                    Case "000001"       '圖面委託
                        URL = "MapSheet_01.aspx?pFormNo=" & wFormNo & _
                                              "&pFormSno=" & wFormSno & _
                                              "&pStep=" & SPDStep & _
                                              "&pSeqNo=" & wSeqNo & _
                                              "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000002"       '圖面修改委託
                        URL = "MapModSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & SPDStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000003"       '內製委託
                        URL = "ManufInSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & SPDStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000004"       '工程連絡單
                        URL = "ManufInOPSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000005"       '業務連絡單
                        URL = "ManufInCTSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000007"       '外注委託
                        URL = "ManufOutSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & SPDStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000008"       '外注工程連絡單
                        URL = "ManufOutOPSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000009"       '外注業務連絡單
                        URL = "ManufOutCTSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000010"       '拉頭細目單
                        URL = "ManufOutSDSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000011"       '進口/客供拉頭
                        URL = "ImportSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000012"       '型別胴體追加委託書
                        URL = "AppendSpecSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000013"       '進口/客供拉頭連絡書
                        URL = "ImportCTSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000014"       '表面處理委託書
                        URL = "SufaceSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "000015"       '表面處理追加委託書
                        URL = "SufaceAppendSheet_01.aspx?pFormNo=" & wFormNo & _
                                                   "&pFormSno=" & wFormSno & _
                                                   "&pStep=" & SPDStep & _
                                                   "&pSeqNo=" & wSeqNo & _
                                                   "&pApplyID=" & Request.QueryString("pUserID")
                    Case "800001"
                        URL = "ModifyDataSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "900050"        '舊圖面委託單
                        URL = "MapBefSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "900051"        '舊內製委託單
                        URL = "ManufInBefSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case "900052"        '舊外注委託單
                        URL = "ManufOutBefSheet_01.aspx?pFormNo=" & wFormNo & _
                                                 "&pFormSno=" & wFormSno & _
                                                 "&pStep=" & wStep & _
                                                 "&pSeqNo=" & wSeqNo & _
                                                 "&pApplyID=" & Request.QueryString("pUserID")
                    Case Else
                End Select

                Response.Redirect(URL)
            End If
        End If
        Response.Write(YKK.ShowMessage(Request.QueryString("pUserID")))
    End Sub

    Private Sub DSimulation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DSimulation.Click
        Dim wFormNo As String = DNewSheet.SelectedValue    '表單號碼
        Dim wFormSno As String = 0       '單號
        Dim wStep As Integer = 1         '工程代碼
        Dim SPDStep As Integer = 1        'SPD工程代碼
        Dim wLevel As String = DLevel.SelectedValue    '難易度
        Dim URL As String = ""            'URL

        If Request.QueryString("pUserID") <> "" Then
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()
            SQL = "Select * From M_Users "
            SQL = SQL & " Where Active =  '1' "
            SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                If DBDataSet1.Tables("M_Users").Rows(0).Item("DIVID") = "1118211" Or _
                   DBDataSet1.Tables("M_Users").Rows(0).Item("DIVID") = "1018950" Then
                    SPDStep = 2
                End If
            End If
            'DB連結關閉
            OleDbConnection1.Close()

            Select Case wFormNo
                Case "000001"       '圖面委託
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000002"       '圖面修改委託
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000003"       '內製委託
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000004"       '工程連絡單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000005"       '業務連絡單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000007"       '外注委託
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000008"       '外注工程連絡單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000009"       '外注業務連絡單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000010"       '業務連絡單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000011"       '進口/客供
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000012"       '型別/胴體
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000013"       '進口/客供業務連絡單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000014"       '表面處理委託單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "000015"       '表面處理追加委託單
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & SPDStep & _
                                          "&pLevel=" & wLevel
                Case "800001"
                    URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                                          "&pFormSno=" & wFormSno & _
                                          "&pStep=" & wStep & _
                                          "&pLevel=" & wLevel
                Case Else
            End Select

            Response.Redirect(URL)
        End If
    End Sub

End Class