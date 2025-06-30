Imports System.Data
Imports System.Data.OleDb

Public Class OPInqCommission_ing
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DCpsc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button

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
    Dim pSts As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPInqCommission_ing.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
            'If pFormNo <> "" And pSts <> "" Then
            'DataList()
            'End If
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
        pSts = Request.QueryString("pSts")
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
        SQL = SQL + "  And FormNo <= '900000' "
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


        '依賴部門
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        '日期
        If pSts = "" Then
            DSDate.Text = CStr(DateAdd("d", -31, DateTime.Now.Today))
            DEDate.Text = CStr(DateTime.Now.Today)
        End If
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue

        Search_Item_Attribute()

    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then
            DataGrid1.Columns.Item(0).HeaderText = "委託No"
            DataGrid1.Columns.Item(1).HeaderText = "依賴者"
            DataGrid1.Columns.Item(2).HeaderText = "依賴日"
            DataGrid1.Columns.Item(3).HeaderText = "完成日"
            SQL = "SELECT "
            'SQL = SQL + wTableName + ".No as Field1, "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
            SQL = SQL + wTableName + ".Division + '--' + "
            SQL = SQL + wTableName + ".Person as Field2, "
            SQL = SQL + "V_WaitHandle_01.ApplyTime as Field3, "
            SQL = SQL + wTableName + ".CompletedTime as Field4, "
            If pFormNo = "000001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "設計者"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000002" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "原圖號"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "設計者"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".OriMapNo as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".MakeMap as Field8, "
            End If

            If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or _
               pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                If pFormNo = "000004" Or pFormNo = "000005" Or _
                   pFormNo = "000008" Or pFormNo = "000009" Then
                    If pFormNo = "000004" Or pFormNo = "000005" Then
                        SQL = SQL + "F_ManufInModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufInModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufInModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufInModSheet.SliderCode as Field8, "
                    Else
                        SQL = SQL + "F_ManufOutModSheet.Buyer   as Field5, "
                        SQL = SQL + "F_ManufOutModSheet.Spec    as Field6, "
                        SQL = SQL + "F_ManufOutModSheet.MapNo   as Field7, "
                        SQL = SQL + "F_ManufOutModSheet.SliderCode as Field8, "
                    End If
                Else
                    SQL = SQL + wTableName + ".Buyer   as Field5, "
                    SQL = SQL + wTableName + ".Spec    as Field6, "
                    SQL = SQL + wTableName + ".MapNo   as Field7, "
                    SQL = SQL + wTableName + ".SliderCode as Field8, "
                End If
            End If

            If pFormNo = "000010" Then
                DataGrid1.Columns.Item(4).HeaderText = "G-Slider Code"
                DataGrid1.Columns.Item(5).HeaderText = "新委託"
                DataGrid1.Columns.Item(6).HeaderText = ""
                DataGrid1.Columns.Item(7).HeaderText = ""

                SQL = SQL + wTableName + ".SliderGRCode as Field5, "
                SQL = SQL + wTableName + ".OFormSno as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If

            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + V_WaitHandle_01.FormNo + "
            SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_01.FormSno,Len(V_WaitHandle_01.FormSno)) + "
            SQL = SQL + "'&pStep='    + str(V_WaitHandle_01.Step,Len(V_WaitHandle_01.Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_01.SeqNo,Len(V_WaitHandle_01.SeqNo)) + "
            SQL = SQL + "'&pApplyID=' + V_WaitHandle_01.ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Left Outer Join V_WaitHandle_01 ON " + wTableName + ".FormNo=V_WaitHandle_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_01.FormSno "
            '內製工程業務連絡
            If pFormNo = "000004" Or pFormNo = "000005" Then
                SQL = SQL + "Left Outer Join F_ManufInModSheet ON " + wTableName + ".NFormNo=F_ManufInModSheet.FormNo "
                SQL = SQL + "                                 And " + wTableName + ".NFormSno=F_ManufInModSheet.FormSno "
            End If
            '外製, 工程業務連絡
            If pFormNo = "000008" Or pFormNo = "000009" Then
                SQL = SQL + "Left Outer Join F_ManufOutModSheet ON " + wTableName + ".NFormNo=F_ManufOutModSheet.FormNo "
                SQL = SQL + "                                  And " + wTableName + ".NFormSno=F_ManufOutModSheet.FormSno "
            End If
            '------------------------------------
            SQL = SQL + "Where V_WaitHandle_01.Step < '10' "
            '表單
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '狀態
            SQL = SQL + " And " + wTableName + ".Sts = '0' "

            '部門
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Division = '" + DDivision.SelectedValue + "'"
            End If
            '依賴人
            If DPerson.Text <> "依賴者" And DPerson.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Person Like '%" + DPerson.Text + "%'"
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Buyer Like '%" + DBuyer.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Buyer Like '%" + DBuyer.Text + "%'"
                    End If
                End If
            End If
            '設計者
            If DMakeMap.Text <> "設計者" And DMakeMap.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".MakeMap Like '%" + DMakeMap.Text + "%'"
            End If
            'Slider Code
            If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.SliderCode Like '%" + DSliderCode.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                    End If
                End If
            End If
            'Spec
            If DSpec.Text <> "Size-Chain-胴體" And DSpec.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Spec Like '%" + DSpec.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Spec Like '%" + DSpec.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Spec Like '%" + DSpec.Text + "%'"
                    End If
                End If
            End If
            'MapNo
            If DMapNo.Text <> "圖號" And DMapNo.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.MapNo Like '%" + DMapNo.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".MapNo Like '%" + DMapNo.Text + "%'"
                    End If
                End If
            End If
            'Cpsc
            If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then
                If pFormNo = "000004" Or pFormNo = "000005" Then
                    SQL = SQL + " And F_ManufInModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                Else
                    If pFormNo = "000008" Or pFormNo = "000009" Then
                        SQL = SQL + " And F_ManufOutModSheet.Cpsc Like '%" + DCpsc.Text + "%'"
                    Else
                        SQL = SQL + " And " + wTableName + ".Cpsc Like '%" + DCpsc.Text + "%'"
                    End If
                End If
            End If
            'No
            If DNo.Text <> "委託單No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub

    Sub Search_Item_Attribute()
        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DMakeMap.ReadOnly = True
        DSpec.ReadOnly = True
        DMapNo.ReadOnly = True
        DSliderCode.ReadOnly = True

        '依賴者
        DPerson.Text = "依賴者"
        'Buyer
        DBuyer.Text = "Buyer"
        '設計者
        DMakeMap.Text = "設計者"
        'SliderCode
        DSliderCode.Text = "Slider Code"
        'Spec
        DSpec.Text = "Size-Chain-胴體"
        'No
        DNo.Text = "委託單No."
        '圖號
        DMapNo.Text = "圖號"
        'Cpsc
        DCpsc.Text = "CPSC"

        DNo.ReadOnly = False
        DPerson.ReadOnly = False
        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
            DMakeMap.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
            DMakeMap.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or + _
           pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
            DBuyer.ReadOnly = False
            DSliderCode.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000010" Then
            DSliderCode.ReadOnly = False
        End If

    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue
        Search_Item_Attribute()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        pSts = DSts.SelectedValue
        DataList()
    End Sub
End Class
