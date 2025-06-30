Imports System.Data
Imports System.Data.OleDb

Public Class DevelopDataPicker
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DKey As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BKey As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Public YKK As New YKK_SPDClass   'YKK共通涵式

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pKey As String     'Search Key

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not Me.IsPostBack Then   '不是PostBack
            If Request.Cookies("CommissionNo").Value <> "" Then
                DKey.Text = Request.Cookies("CommissionNo").Value
                pKey = Request.Cookies("CommissionNo").Value
            Else
                DKey.Text = ""
                pKey = "ALL"
            End If
            CodeData()  '取得資料
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandSource.CommandName = "Select" Then  '點選選取檢查
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '所選取的Code

            Dim SQL, Str As String
            Dim wRNO, wAPPBuyer, wSizeNo, wItem, wCodeNo, wTAWidth, wDevNo, wDevPrd, wTACol, wTALine, wECol, wCCol, wTHCol As String
            Dim wTNLItem, wTNRItem, wTSLItem, wTSRItem, wTDLItem, wTDRItem, wCNItem, wCSItem, wCDItem, wCItem, wOP1, wOP2, wOP3, wOP4, wOP5, wOP6, wOther1, wOther2, wO1Item, wO2Item As String

            Dim DBDataSet1 As New DataSet
            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SCDSqlConn")  'SQL連結設定
            OleDbConnection1.Open()

            SQL = "Select * From V_SampleFile_01 "
            SQL = SQL & " Where RNO = '" & Key & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "SampleFile")
            If DBDataSet1.Tables("SampleFile").Rows.Count > 0 Then
                wRno = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("Rno")))
                wAPPBuyer = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("APPBuyer")))
                wSizeNo = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("SizeNo")))
                wItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("Item")))
                wCodeNo = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CodeNo")))
                wTAWidth = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TAWidth")))
                wDevNo = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("DevNo")))
                '--期間--------------------------------------------------------------
                Str = ""
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("ReDate") <> "1900/1/1" Then Str = DBDataSet1.Tables("SampleFile").Rows(0).Item("ReDate")
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("DevCDate") <> "1900/1/1" Then Str = Str + "~" + DBDataSet1.Tables("SampleFile").Rows(0).Item("DevCDate")
                wDevPrd = Str
                '--布帶--------------------------------------------------------------
                Str = ""
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TACol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TAColNo"))) <> "" Then
                    If Str <> "" Then Str = Str + ", "
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TACol"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TAColNo")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TALCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TALColNo"))) <> "" Then
                    If Str <> "" Then Str = Str + ", "
                    Str = Str + "TALCOL" + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TALCol"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TALColNo")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TARCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TARColNo"))) <> "" Then
                    If Str <> "" Then Str = Str + ", "
                    Str = Str + "TARCOL" + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TARCol"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TARColNo")))
                End If
                wTACol = Str
                '--條紋線--------------------------------------------------------------
                Str = ""
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("XMLen") <> 0 Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("XMCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("XMColNo"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", X:"
                    Else
                        Str = Str + "X:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(CStr(DBDataSet1.Tables("SampleFile").Rows(0).Item("XMLen")))) + "mm " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("XMCol"))) + " " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("XMColNo")))
                End If
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("AMLen") <> 0 Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("AMCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("AMColNo"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", A:"
                    Else
                        Str = Str + "A:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(CStr(DBDataSet1.Tables("SampleFile").Rows(0).Item("AMLen")))) + "mm " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("AMCol"))) + " " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("AMColNo")))
                End If
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("BMLen") <> 0 Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("BMCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("BMColNo"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", B:"
                    Else
                        Str = Str + "B:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(CStr(DBDataSet1.Tables("SampleFile").Rows(0).Item("BMLen")))) + "mm " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("BMCol"))) + " " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("BMColNo")))
                End If
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("CMLen") <> 0 Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CMCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CMColNo"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", C:"
                    Else
                        Str = Str + "C:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(CStr(DBDataSet1.Tables("SampleFile").Rows(0).Item("CMLen")))) + "mm " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CMCol"))) + " " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CMColNo")))
                End If
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("DMLen") <> 0 Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("DMCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("DMColNo"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", D:"
                    Else
                        Str = Str + "D:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(CStr(DBDataSet1.Tables("SampleFile").Rows(0).Item("DMLen")))) + "mm " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("DMCol"))) + " " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("DMColNo")))
                End If
                If DBDataSet1.Tables("SampleFile").Rows(0).Item("FMLen") <> 0 Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("FMCol"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("FMColNo"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", F:"
                    Else
                        Str = Str + "F:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(CStr(DBDataSet1.Tables("SampleFile").Rows(0).Item("FMLen")))) + "mm " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("FMCol"))) + " " + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("FMColNo")))
                End If
                wTALine = Str
                '----------------------------------------------------------------
                wECol = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("ECol")))
                wCCol = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CCol")))
                '--縫工線--------------------------------------------------------------
                Str = ""
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THUPCOL"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THUPCOLNO"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 上:"
                    Else
                        Str = Str + "上:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THUPCOL"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THUPCOLNO")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLUPCOL"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLUPCOLNO"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 上左:"
                    Else
                        Str = Str + "上左:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLUPCOL"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLUPCOLNO")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRUPCOL"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRUPCOLNO"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 上右:"
                    Else
                        Str = Str + "上右:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRUPCOL"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRUPCOLNO")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLOCOL"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLOCOLNO"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 下:"
                    Else
                        Str = Str + "下:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLOCOL"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLOCOLNO")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLLOCOL"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLLOCOLNO"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 下左:"
                    Else
                        Str = Str + "下左:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLLOCOL"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THLLOCOLNO")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRLOCOL"))) <> "" Or _
                   LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRLOCOLNO"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 下右:"
                    Else
                        Str = Str + "下右:"
                    End If
                    Str = Str + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRLOCOL"))) + _
                          LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("THRLOCOLNO")))
                End If
                wTHCol = Str
                '----------------------------------------------------------------
                wTNLItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TNLItem")))
                wTNRItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TNRItem")))
                wTSLItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TSLItem")))
                wTSRItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TSRItem")))
                wTDLItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TDLItem")))
                wTDRItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TDRItem")))
                wCNItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CNItem")))
                wCSItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CSItem")))
                wCDItem = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("CDItem")))
                '----------------------------------------------------------------
                Str = ""
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TCLItem"))) <> "" Then
                    Str = "左:" + LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TCLItem")))
                End If
                If LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TCRItem"))) <> "" Then
                    If Str <> "" Then
                        Str = Str + ", 右:"
                    Else
                        Str = Str + "右:"
                    End If
                    Str = Str + LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("TCRItem")))
                End If
                wCItem = Str
                '----------------------------------------------------------------
                wOP1 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("OP1")))
                wOP2 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("OP2")))
                wOP3 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("OP3")))
                wOP4 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("OP4")))
                wOP5 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("OP5")))
                wOP6 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("OP6")))

                wO1Item = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("O1Item")))
                wO2Item = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("O2Item")))
                wOther1 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("Other1")))
                wOther2 = LTrim(RTrim(DBDataSet1.Tables("SampleFile").Rows(0).Item("Other2")))
            End If
            'DB連結關閉
            OleDbConnection1.Close()

            '所選取的Data傳回父視窗並關閉
            Dim Cmd As String
            Response.Cookies("CommissionNo").Value = Key

            Cmd = "<script>" + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DRNO", wRNO) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DAPPBUYER", wAPPBuyer) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DSIZENO", wSizeNo) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DITEM", wItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DCODENO", wCodeNo) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTAWIDTH", wTAWidth) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DDEVNO", wDevNo) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DDEVPRD", wDevPrd) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTACOL", wTACol) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTALINE", wTALine) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DECOL", wECol) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DCCOL", wCCol) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTHCOL", wTHCol) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTNLITEM", wTNLItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTNRITEM", wTNRItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTSLITEM", wTSLItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTSRITEM", wTSRItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTDLITEM", wTDLItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DTDRITEM", wTDRItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DCNITEM", wCNItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DCSITEM", wCSItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DCDITEM", wCDItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DCITEM", wCItem) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DOP1", wOP1) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DOP2", wOP2) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DOP3", wOP3) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DOP4", wOP4) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DOP5", wOP5) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DOP6", wOP6) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DO1ITEM", wO1Item) + _
                  String.Format("window.opener.document.{0}.value = '{1:d}';", "Form1.DO2ITEM", wO2Item) + _
                  String.Format("window.opener.document.getElementById('{0}').value = '{1:d}';", "Hidden1", wOther1) + _
                  String.Format("window.opener.document.getElementById('{0}').value = '{1:d}';", "Hidden2", wOther2) + _
                  String.Format("window.opener.document.getElementById('{0}').innerHTML = '{1:d}';", "D1Other", wOther1) + _
                  String.Format("window.opener.document.getElementById('{0}').innerHTML = '{1:d}';", "D2Other", wOther2) + _
                    "window.close();" + _
                  "</script>"

            'Response.Write(YKK.ShowMessage(wTACol))
            Response.Write(Cmd)

        End If
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        'DataGrid1.DataBind()
        CodeData()
    End Sub

    Sub CodeData()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SCDSqlConn")  'SQL連結設定
        Dim SQL As String

        SQL = "Select Rno, CodeNo, DevNo From V_SampleFile_01 "   '取得DB資料
        SQL = SQL & " Where (Sts = '6' or Sts = '8') "
        If (pKey <> "ALL") And (pKey <> "") Then
            SQL = SQL + "And (RNo Like '%" & pKey & "%' or CodeNo Like '%" & pKey & "%' or DevNo Like '%" & pKey & "%') "
        End If
        SQL = SQL + "Order by CRDate Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "SampleFile")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub BKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BKey.Click
        DataGrid1.CurrentPageIndex = 0  'DataGrid跳上下頁
        pKey = DKey.Text
        CodeData()
    End Sub

End Class
