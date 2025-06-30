Imports System.Data
Imports System.Data.OleDb

Public Class GetSeqNo_T
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents DSeqno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DBody As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DChainType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Dsize As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents Button4 As System.Web.UI.WebControls.Button
    Protected WithEvents DQAContent As System.Web.UI.WebControls.ListBox
    Protected WithEvents DQAContent1 As System.Web.UI.WebControls.ListBox
    Protected WithEvents Button5 As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        Button3.Attributes("onclick") = "QAPicker('Form1.DQAContent');"
        Button4.Attributes("onclick") = "QA2Picker('Form1.DQAContent1');"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim oGetSeqNo As Object
        Dim pFormNo
        Dim pSeqNo
        Dim RtnCode

        pFormNo = DFormNo.Text
        pSeqNo = 0
        oGetSeqNo = Server.CreateObject("GetSeqno.WFFormInf")
        RtnCode = oGetSeqNo.Seqno(pFormNo, pSeqNo)

        If RtnCode = 1 Then
            DSeqno.Text = "Error"
        Else
            DSeqno.Text = CStr(pSeqNo)
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim SQL As String
        Dim TableName As String = ""
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定


        SQL = "SELECT TableName1 FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo = '" & "000001" & "' "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FORM")
        If DBDataSet1.Tables("M_FORM").Rows.Count > 0 Then
            TableName = "V_" & DBDataSet1.Tables("M_FORM").Rows(0).Item("TableName1") & "_01"
        End If
        OleDbConnection1.Close()

        DBDataSet1.Clear()

        'If TableName <> "" Then
        SQL = "SELECT * from " & TableName
        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "Sheet")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()
        'End If
    End Sub

    Private Sub BSpec_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSpec.Click
        If DSpec.Text = "" Then
            DSpec.Text = Dsize.SelectedValue & "," & DChainType.SelectedValue & "," & DBody.SelectedValue
        Else
            DSpec.Text = DSpec.Text & ":" & Dsize.SelectedValue & "," & DChainType.SelectedValue & "," & DBody.SelectedValue
        End If

    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString += "Show('aaa');"
        scriptString += "</script>"

        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If

    End Sub
End Class
