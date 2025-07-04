Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class MultiOPPerformance
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DSStep As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DEStep As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList

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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "MultiOPPerformance.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            pFormNo = "000001"
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時

        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim wSts As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '表單
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000') or FormNo='800001' ) "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        'Step
        DBDataSet1.Clear()
        SQL = "Select Step, str(Step) + ':' + StepName As StepName From M_Flow "
        SQL = SQL + "Where Step     <  '900' "
        SQL = SQL + "  and FormNo   =  '" + pFormNo + "' "
        SQL = SQL + "  and FlowType <> '0'   "
        SQL = SQL + "Group by Step, StepName "
        SQL = SQL + "Order by Step, StepName "
        DSStep.Items.Clear()
        DEStep.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        DBTable1 = DBDataSet1.Tables("M_Flow")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("StepName")
            ListItem1.Value = DBTable1.Rows(i).Item("Step")
            DSStep.Items.Add(ListItem1)
            DEStep.Items.Add(ListItem1)
        Next

        '依賴部門
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        '日期
        DSDate.Text = CStr(DateTime.Now.Year) & "/" & CStr(DateTime.Now.Month) & "/1"
        DEDate.Text = CStr(DateTime.Now.Today)

        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue
        Search_Item_Attribute()

    End Sub

    Sub Search_Item_Attribute()
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True

        '依賴者
        DPerson.Text = "依賴者"
        'Buyer
        DBuyer.Text = "Buyer"

        DPerson.ReadOnly = False
        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or + _
           pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000010" Then
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        SetSearchItem()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        pFormNo = DFormName.SelectedValue
        CreateWorkTable()
    End Sub

    Sub CreateWorkTable()
        Dim i As Integer
        Dim SQL As String
        Dim wTableName As String = ""
        Dim wFormNo As String = ""
        Dim wFormSno As Integer = 0
        Dim wStep As Integer = 0

        Dim wSStep As Integer = 0
        Dim wEStep As Integer = 0
        Dim wSDate As String = DSDate.Text & " 00:00:00"
        Dim wEDate As String = DEDate.Text & " 23:59:59"
        Dim WorkTable As String = "Temp_MultiOPPerformance_" & Request.Cookies("UserID").Value

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBDataSet3 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()

        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Form")
        If DBDataSet1.Tables("Form").Rows.Count > 0 Then
            wTableName = "V_" + DBDataSet1.Tables("Form").Rows(0).Item("TableName1") + "_01"
        End If

        DBDataSet1.Clear()

        SQL = "SELECT FormNo, FormSno, No, Sts, Division, BWorkTime as BTime, AWorkTime as ATime, T_AStartTime as AStartTime, T_AEndTime as AEndTime, "
        SQL = SQL + "T_Step as Step, T_SeqNo as SeqNo, T_CompletedTime, WorkID, "
        If DFormName.SelectedValue = "000001" Then
            SQL = SQL + "Buyer, "
            SQL = SQL + "Case HalfFinish When 'Yes' Then '半成品' Else '成品' End As Finished, "
            SQL = SQL + "Suppiler As Suppiler, "
            SQL = SQL + "Level    As Level "
        End If
        If DFormName.SelectedValue = "000002" Then
            SQL = SQL + "Buyer, "
            SQL = SQL + "'' As Finished, "
            SQL = SQL + "Suppiler As Suppiler, "
            SQL = SQL + "Level    As Level "
        End If
        If DFormName.SelectedValue = "000003" Or DFormName.SelectedValue = "000007" Then
            SQL = SQL + "Buyer, "
            SQL = SQL + "SliderType2 As Finished, "
            SQL = SQL + "Suppiler As Suppiler, "
            SQL = SQL + "Level    As Level "
        End If
        If DFormName.SelectedValue = "000005" Or DFormName.SelectedValue = "000009" Or DFormName.SelectedValue = "000010" Or DFormName.SelectedValue = "000013" Then
            SQL = SQL + "'' As Buyer, "
            SQL = SQL + "'' As Finished, "
            SQL = SQL + "'' As Suppiler, "
            SQL = SQL + "'' As Level "
        End If
        If DFormName.SelectedValue = "000011" Or DFormName.SelectedValue = "000012" Then
            SQL = SQL + "Buyer, "
            SQL = SQL + "'' As Finished, "
            SQL = SQL + "'' As Suppiler, "
            SQL = SQL + "'' As Level "
        End If
        SQL = SQL + "FROM " + wTableName + " "
        SQL = SQL + "Where T_Active = '0' "
        SQL = SQL + "  And T_FormNo = '" + DFormName.SelectedValue + "' "
        SQL = SQL + "  And T_Step   >= '" + DSStep.SelectedValue + "' "
        SQL = SQL + "  And T_Step   <= '" + DEStep.SelectedValue + "' "
        SQL = SQL + "  And T_CompletedTime >= '" + wSDate + "' "
        SQL = SQL + "  And T_CompletedTime <= '" + wEDate + "' "
        SQL = SQL + "  And FormSno  > '2000' "
        SQL = SQL + "  And Action = '0' "
        SQL = SQL + "  And T_FlowType <> '0' "

        If DSts.SelectedValue <> "ALL" Then
            SQL = SQL + " And Sts = '" + DSts.SelectedValue + "'"
        End If
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
        End If
        If DPerson.Text <> "依賴者" And DPerson.Text <> "" Then
            SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
        End If
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
            SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
        End If

        SQL = SQL + "Order by FormNo, FormSno, T_Step, AWorkTime Desc "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Sheet")
        DBTable1 = DBDataSet1.Tables("Sheet")

        If DBDataSet1.Tables("Sheet").Rows.Count > 0 Then
            'Call Stored Procedure Create WorkTable
            SQL = "Exec sp_Temp_MultiOPPerformance '" & WorkTable & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        End If

        For i = 0 To DBTable1.Rows.Count - 1
            If wFormNo <> DBDataSet1.Tables("Sheet").Rows(i).Item("FormNo") Or _
               wFormSno <> DBDataSet1.Tables("Sheet").Rows(i).Item("FormSno") Then

                DBDataSet2.Clear()
                SQL = "SELECT Step FROM T_WaitHandle "
                SQL = SQL + "Where FormNo = '" + DBDataSet1.Tables("Sheet").Rows(i).Item("FormNo") + "'"
                SQL = SQL + "  And FormSno = '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("FormSno")) + "'"
                SQL = SQL + "  And Step   >= '" + DSStep.SelectedValue + "' "
                SQL = SQL + "  And Step   <= '" + DEStep.SelectedValue + "' "
                SQL = SQL + "  And CompletedTime >= '" + wSDate + "' "
                SQL = SQL + "  And CompletedTime <= '" + wEDate + "' "
                SQL = SQL + "  And FlowType <> '0' "
                SQL = SQL + "Order by Step "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet2, "WaitHandle")
                If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                    wSStep = DBDataSet2.Tables("WaitHandle").Rows(0).Item("Step")
                End If

                DBDataSet2.Clear()
                SQL = "SELECT Step FROM T_WaitHandle "
                SQL = SQL + "Where FormNo = '" + DBDataSet1.Tables("Sheet").Rows(i).Item("FormNo") + "'"
                SQL = SQL + "  And FormSno = '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("FormSno")) + "'"
                SQL = SQL + "  And Step   >= '" + DSStep.SelectedValue + "' "
                SQL = SQL + "  And Step   <= '" + DEStep.SelectedValue + "' "
                SQL = SQL + "  And CompletedTime >= '" + wSDate + "' "
                SQL = SQL + "  And CompletedTime <= '" + wEDate + "' "
                SQL = SQL + "  And FlowType <> '0' "
                SQL = SQL + "Order by Step Desc "
                Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter4.Fill(DBDataSet2, "WaitHandle")
                If DBDataSet2.Tables("WaitHandle").Rows.Count > 0 Then
                    wEStep = DBDataSet2.Tables("WaitHandle").Rows(0).Item("Step")
                End If

                wFormNo = DBDataSet1.Tables("Sheet").Rows(i).Item("FormNo")
                wFormSno = DBDataSet1.Tables("Sheet").Rows(i).Item("FormSno")
            End If

            If wStep <> DBDataSet1.Tables("Sheet").Rows(i).Item("Step") Then
                SQL = "Insert into " & WorkTable & " "
                SQL = SQL + "( "
                SQL = SQL + "RecordType, FormNo, FormSno, Step, SeqNo, SStep, EStep, CompletedTime, "     '1~7
                SQL = SQL + "Buyer, No, Sts, Division, Finished, Suppiler, Level, BTime, ATime, TimeRange, WorkID, "  '8~16
                SQL = SQL + "CreateUser, CreateTime "  '17~18
                SQL = SQL + ")  "
                SQL = SQL + "VALUES( "
                SQL = SQL + " '0', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("FormNo") + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("FormSno")) + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("Step")) + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("SeqNo")) + "', "
                SQL = SQL + " '" + CStr(wSStep) + "', "
                SQL = SQL + " '" + CStr(wEStep) + "', "

                Dim Str As String = CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("T_CompletedTime").Date) + " " + _
                                    CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("T_CompletedTime").Hour) + ":" + _
                                    CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("T_CompletedTime").Minute) + ":" + _
                                    CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("T_CompletedTime").Second)
                SQL = SQL + " '" + Str + "', "

                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Buyer") + "', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("No") + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("Sts")) + "', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Division") + "', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Finished") + "', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Suppiler") + "', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Level") + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("BTime")) + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("ATime")) + "', "
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AStartTime")) + " ~ " + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AEndTime")) + "', "
                SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("WorkID") + "', "

                SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '作成者
                SQL = SQL + " '" + NowDateTime + "' "       '作成時間
                SQL = SQL + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQL
                OleDBCommand1.ExecuteNonQuery()

                wStep = DBDataSet1.Tables("Sheet").Rows(i).Item("Step")
            End If
        Next

        DBDataSet1.Clear()
        DBDataSet2.Clear()

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, SStep, EStep, Buyer, No, Sts, Division, "
        SQL = SQL + "Finished, Suppiler, Level,  Sum(BTime) As BTime, Sum(ATime)  As ATime "
        SQL = SQL + "FROM " & WorkTable & " "
        SQL = SQL + "group by FormNo, FormSno, SStep, EStep, Buyer, No, Sts, Division, Finished, Suppiler, Level "
        SQL = SQL + "order  by FormNo, FormSno, SStep, EStep, Buyer, No, Sts, Division, Finished, Suppiler, Level "
        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet1, "Group")
        DBTable1 = DBDataSet1.Tables("Group")

        For i = 0 To DBTable1.Rows.Count - 1
            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "RecordType, FormNo, FormSno, Step, SStep, EStep, CompletedTime, "     '1~7
            SQL = SQL + "Buyer, No, Sts, Division, Finished, Suppiler, Level, BTime, ATime, TimeRange, "  '8~16
            SQL = SQL + "CreateUser, CreateTime "  '17~18
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '1', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("FormNo") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("FormSno")) + "', "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("SStep")) + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("EStep")) + "', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Buyer") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("No") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("Sts")) + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Division") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Finished") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Suppiler") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Level") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("BTime")) + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("ATime")) + "', "
            SQL = SQL + " '" + "" + "', "
            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '作成者
            SQL = SQL + " '" + NowDateTime + "' "       '作成時間
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next

        SQL = "SELECT No, FormNo, FormSno, SStep, EStep, Division, Buyer, Finished, Suppiler, Level, BTime, ATime, "
        SQL = SQL + "str(SStep)+' ~ '+ str(EStep) As Step, "
        SQL = SQL + "Case Sts When 0 Then '開發中' When 1 Then 'OK' When 2 Then 'NG' Else '取消' End As Sts, "
        SQL = SQL + "BTime-ATime As Balance, "

        SQL = SQL + "'OPPerformanceList.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL = SQL + "'&pTable=' + '" & WorkTable & "'"
        SQL = SQL + " As URL "

        SQL = SQL + "FROM " + WorkTable + " "
        SQL = SQL + "Where RecordType = '1' "
        SQL = SQL + "Order by No "

        Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter6.Fill(DBDataSet3, "Header")
        DataGrid1.DataSource = DBDataSet3
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

End Class
