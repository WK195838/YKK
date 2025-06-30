Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class OPInqCommission_List
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DCpsc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
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
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DProgress As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DMPSelect As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKeepData As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSUPPILER As System.Web.UI.WebControls.TextBox
    Protected WithEvents LISIP As System.Web.UI.WebControls.HyperLink

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
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim SaveTime As String
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPInqCommission_List.aspx"

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
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
        '
        'DB連結設定
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        ' SQL = "Select * From M_Referp Where Cat='018' and Dkey = 'ALL'  Order by DKey, Data "
        SQL = " select substring(convert(char(10),dateadd(year,-3,getdate()),111),1,8)+'01' as Data "


        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            SaveTime = DBDataSet1.Tables("M_Referp").Rows(0).Item("Data")
        Else
            SaveTime = ""
        End If

        'DB連結關閉
        OleDbConnection1.Close()


        LISIP.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & Request.QueryString("pUserID")

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
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000') or FormNo='800001' or FormNo= '007001' ) "
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
        DSDate.Text = CStr(DateAdd("d", -180, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue

        Search_Item_Attribute()

    End Sub
    '****************************************************
    '  未封存資料
    '****************************************************
    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
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

            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                DataGrid1.Columns.Item(0).HeaderText = "委託No"
                DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
                DataGrid1.Columns.Item(2).HeaderText = "加工種類"
                DataGrid1.Columns.Item(3).HeaderText = "外注廠商"

                SQL = "SELECT "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "

                SQL = SQL + wTableName + ".Material as Field3, "
                SQL = SQL + wTableName + ".Suppiler as Field4, "
            Else
                DataGrid1.Columns.Item(0).HeaderText = "委託No"
                DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
                DataGrid1.Columns.Item(2).HeaderText = "依賴者"
                DataGrid1.Columns.Item(3).HeaderText = "依賴日"
                SQL = "SELECT "
                'SQL = SQL + wTableName + ".No as Field1, "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "

                SQL = SQL + wTableName + ".Division + '--' + "
                SQL = SQL + wTableName + ".Person as Field3, "


                SQL = SQL + "V_WaitHandle_01.ApplyTime as Field4, "
                'SQL = SQL + wTableName + ".CompletedTime as Field4, "
            End If

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
            If pFormNo = "000011" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + "'  ' as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "OR-No"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".ORNO    as Field8, "
            End If
            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "追加理由"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".AppendReason    as Field8, "
            End If
            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "色番項目"
                DataGrid1.Columns.Item(5).HeaderText = "購入價/加價"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".ColorItem  as Field5, "
                SQL = SQL + wTableName + ".Pullerprice" + " +'/'+ " + wTableName + ".price as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "

            End If
            If pFormNo = "800001" Then
                DataGrid1.Columns.Item(4).HeaderText = "委託書"
                DataGrid1.Columns.Item(5).HeaderText = "狀態"
                DataGrid1.Columns.Item(6).HeaderText = "修改理由類別"
                DataGrid1.Columns.Item(7).HeaderText = "修改理由"

                SQL = SQL + wTableName + ".Sheet   as Field5, "
                SQL = SQL + wTableName + ".Status  as Field6, "
                SQL = SQL + wTableName + ".MReasonType as Field7, "
                SQL = SQL + wTableName + ".MReason as Field8, "
            End If

            If pFormNo = "007001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "圖號"
                DataGrid1.Columns.Item(6).HeaderText = "拉頭品名"
                DataGrid1.Columns.Item(7).HeaderText = "姐妹社"


                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".MapNo  as Field6, "
                SQL = SQL + wTableName + ".SliderCode as Field7, "
                SQL = SQL + wTableName + ".YKKGroup as Field8, "
            End If

            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                Dim bc1 As New BoundColumn
                Dim bc2 As New BoundColumn

                bc1.DataField = "Field9"
                bc1.HeaderText = "應收模具費"
                bc2.DataField = "Field10"
                bc2.HeaderText = "模具購入價"

                DataGrid1.Columns.AddAt(8, bc1) ''新增DataGrid欄位
                DataGrid1.Columns.AddAt(9, bc2)
                SetFieldWidth()

                SQL = SQL + wTableName + ".ArMoldFee as Field9, "
                SQL = SQL + wTableName + ".PurMold as Field10, "
                Dim bc3 As New BoundColumn
                bc3.DataField = "Field11"
                bc3.HeaderText = "委託廠商"
                DataGrid1.Columns.AddAt(10, bc3) ''新增DataGrid欄位
                SQL = SQL + wTableName + ".SellVendor as Field11, "
                If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                    Dim bc4 As New BoundColumn
                    bc4.DataField = "Field12"
                    bc4.HeaderText = "外注廠商"
                    DataGrid1.Columns.AddAt(11, bc4) ''新增DataGrid欄位
                    SQL = SQL + wTableName + ".SUPPILER as Field12, "
                End If


            End If
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                SQL = SQL + wTableName + ".MMSSts as MSts, "
            End If

            '20180821 jessica 
            If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
                If DMPSelect.SelectedValue <> "1" Then
                    Dim bc3 As New BoundColumn
                    bc3.DataField = "Field9"
                    bc3.HeaderText = "委託廠商"
                    DataGrid1.Columns.AddAt(8, bc3) ''新增DataGrid欄位
                    SQL = SQL + wTableName + ".SellVendor as Field9, "
                    If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                        Dim bc4 As New BoundColumn
                        bc4.DataField = "Field10"
                        bc4.HeaderText = "外注廠商"
                        DataGrid1.Columns.AddAt(9, bc4) ''新增DataGrid欄位
                        SQL = SQL + wTableName + ".SUPPILER as Field10, "
                    End If

                End If
            End If
            'Add-End






            'Add-End
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
            SQL = SQL + "Where V_WaitHandle_01.Step  < '10' "
            '表單
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '開發狀態
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If

            '開發完成狀態
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If

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

            '20180821 jessica 增加委託廠商
            'sellvendor
            If DSellVendor.Text <> "委託廠商" And DSellVendor.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SellVendor Like '%" + DSellVendor.Text + "%'"

            End If

            '202401191 jessica 增加外注 
            'sellvendor
            If DSUPPILER.Text <> "外注廠商" And DSUPPILER.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SUPPILER Like '%" + DSUPPILER.Text + "%'"

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
                        If pFormNo = "000010" Then
                            SQL = SQL + " And " + wTableName + ".SliderGRCode Like '%" + DSliderCode.Text + "%'"
                        Else
                            SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                        End If
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
            'Code
            If DCode.Text <> "Code" And DCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Code Like '%" + DCode.Text + "%'"
            End If
            ''CompletedTime
            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            Else
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            '
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                DBTable1 = DBDataSet2.Tables("WaitHandle")
                If DBTable1.Rows.Count > 0 Then
                    For i = 0 To DBTable1.Rows.Count - 1
                        If DBTable1.Rows(i).Item("MSts") = 1 Then
                            DBTable1.Rows(i).Item("Field1") = DBTable1.Rows(i).Item("Field1") + "-模廢"
                        End If
                    Next
                End If
            End If
            'Add-End
            '
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub
    '****************************************************
    '  已封存資料
    '****************************************************
    Sub KeepDataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
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

            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                DataGrid1.Columns.Item(0).HeaderText = "委託No"
                DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
                DataGrid1.Columns.Item(2).HeaderText = "加工種類"
                DataGrid1.Columns.Item(3).HeaderText = "外注廠商"

                SQL = "SELECT "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "

                SQL = SQL + wTableName + ".Material as Field3, "
                SQL = SQL + wTableName + ".Suppiler as Field4, "
            Else
                DataGrid1.Columns.Item(0).HeaderText = "委託No"
                DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
                DataGrid1.Columns.Item(2).HeaderText = "依賴者"
                DataGrid1.Columns.Item(3).HeaderText = "依賴日"
                SQL = "SELECT "
                'SQL = SQL + wTableName + ".No as Field1, "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "

                SQL = SQL + wTableName + ".Division + '--' + "
                SQL = SQL + wTableName + ".Person as Field3, "
                SQL = SQL + "V_WaitHandle_OLD_01.ApplyTime as Field4, "
                'SQL = SQL + wTableName + ".CompletedTime as Field4, "
            End If

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
            If pFormNo = "000011" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + "'  ' as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "OR-No"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".ORNO    as Field8, "
            End If
            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "追加理由"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".AppendReason    as Field8, "
            End If
            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "色番項目"
                DataGrid1.Columns.Item(5).HeaderText = "購入價/加價"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".ColorItem  as Field5, "
                SQL = SQL + wTableName + ".Pullerprice" + " +'/'+ " + wTableName + ".price as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "800001" Then
                DataGrid1.Columns.Item(4).HeaderText = "委託書"
                DataGrid1.Columns.Item(5).HeaderText = "狀態"
                DataGrid1.Columns.Item(6).HeaderText = "修改理由類別"
                DataGrid1.Columns.Item(7).HeaderText = "修改理由"

                SQL = SQL + wTableName + ".Sheet   as Field5, "
                SQL = SQL + wTableName + ".Status  as Field6, "
                SQL = SQL + wTableName + ".MReasonType as Field7, "
                SQL = SQL + wTableName + ".MReason as Field8, "
            End If

            If pFormNo = "007001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "圖號"
                DataGrid1.Columns.Item(6).HeaderText = "拉頭品名"
                DataGrid1.Columns.Item(7).HeaderText = "姐妹社"


                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".MapNo  as Field6, "
                SQL = SQL + wTableName + ".SliderCode as Field7, "
                SQL = SQL + wTableName + ".YKKGroup as Field8, "
            End If


            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                Dim bc1 As New BoundColumn
                Dim bc2 As New BoundColumn

                bc1.DataField = "Field9"
                bc1.HeaderText = "應收模具費"
                bc2.DataField = "Field10"
                bc2.HeaderText = "模具購入價"

                DataGrid1.Columns.AddAt(8, bc1) ''新增DataGrid欄位
                DataGrid1.Columns.AddAt(9, bc2)
                SetFieldWidth()

                SQL = SQL + wTableName + ".ArMoldFee as Field9, "
                SQL = SQL + wTableName + ".PurMold as Field10, "
                Dim bc3 As New BoundColumn
                bc3.DataField = "Field11"
                bc3.HeaderText = "委託廠商"
                DataGrid1.Columns.AddAt(10, bc3) ''新增DataGrid欄位
                SQL = SQL + wTableName + ".SellVendor as Field11, "

                If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                    Dim bc4 As New BoundColumn
                    bc4.DataField = "Field12"
                    bc4.HeaderText = "外注廠商"
                    DataGrid1.Columns.AddAt(11, bc4) ''新增DataGrid欄位
                    SQL = SQL + wTableName + ".Suppiler as Field12, "
                End If

            End If
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                SQL = SQL + wTableName + ".MMSSts as MSts, "
            End If

            '20180821 jessica 
            If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
                If DMPSelect.SelectedValue <> "1" Then
                    Dim bc3 As New BoundColumn
                    bc3.DataField = "Field9"
                    bc3.HeaderText = "委託廠商"
                    DataGrid1.Columns.AddAt(8, bc3) ''新增DataGrid欄位
                    SQL = SQL + wTableName + ".SellVendor as Field9, "

                    If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                        Dim bc4 As New BoundColumn
                        bc4.DataField = "Field10"
                        bc4.HeaderText = "外注廠商"
                        DataGrid1.Columns.AddAt(9, bc4) ''新增DataGrid欄位
                        SQL = SQL + wTableName + ".SUPPILER as Field10, "

                    End If


                End If



            End If

            'Add-End

            'Add-End
            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + V_WaitHandle_OLD_01.FormNo + "
            SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_OLD_01.FormSno,Len(V_WaitHandle_OLD_01.FormSno)) + "
            SQL = SQL + "'&pStep='    + str(V_WaitHandle_OLD_01.Step,Len(V_WaitHandle_OLD_01.Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_OLD_01.SeqNo,Len(V_WaitHandle_OLD_01.SeqNo)) + "
            SQL = SQL + "'&pKeepdata='   + rtrim(ltrim(str(1))) + "
            SQL = SQL + "'&pApplyID=' + V_WaitHandle_OLD_01.ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Left Outer Join V_WaitHandle_OLD_01 ON " + wTableName + ".FormNo=V_WaitHandle_OLD_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_OLD_01.FormSno "
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
            SQL = SQL + "Where V_WaitHandle_OLD_01.Step  < '10' "
            '表單
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '開發狀態
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If

            '開發完成狀態
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If

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


            '20180821 jessica 增加委託廠商
            'sellvendor
            If DSellVendor.Text <> "委託廠商" And DSellVendor.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SellVendor Like '%" + DSellVendor.Text + "%'"

            End If

            '202401191 jessica 增加外注 
            'sellvendor
            If DSUPPILER.Text <> "外注廠商" And DSUPPILER.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SUPPILER Like '%" + DSUPPILER.Text + "%'"

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
                        If pFormNo = "000010" Then
                            SQL = SQL + " And " + wTableName + ".SliderGRCode Like '%" + DSliderCode.Text + "%'"
                        Else
                            SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                        End If
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
            'Code
            If DCode.Text <> "Code" And DCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Code Like '%" + DCode.Text + "%'"
            End If
            ''CompletedTime
            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            Else
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            '
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                DBTable1 = DBDataSet2.Tables("WaitHandle")
                If DBTable1.Rows.Count > 0 Then
                    For i = 0 To DBTable1.Rows.Count - 1
                        If DBTable1.Rows(i).Item("MSts") = 1 Then
                            DBTable1.Rows(i).Item("Field1") = DBTable1.Rows(i).Item("Field1") + "-模廢"
                        End If
                    Next
                End If
            End If
            'Add-End
            '
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub
    '****************************************************
    '  限表單
    '****************************************************
    Sub FormDataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
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

            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                DataGrid1.Columns.Item(0).HeaderText = "委託No"
                DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
                DataGrid1.Columns.Item(2).HeaderText = "加工種類"
                DataGrid1.Columns.Item(3).HeaderText = "外注廠商"

                SQL = "SELECT "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "

                SQL = SQL + wTableName + ".Material as Field3, "
                SQL = SQL + wTableName + ".Suppiler as Field4, "
            Else
                DataGrid1.Columns.Item(0).HeaderText = "委託No"
                DataGrid1.Columns.Item(1).HeaderText = "開發狀態"
                DataGrid1.Columns.Item(2).HeaderText = "依賴者"
                DataGrid1.Columns.Item(3).HeaderText = "依賴日"
                SQL = "SELECT "
                'SQL = SQL + wTableName + ".No as Field1, "

                SQL = SQL + "Case " + wTableName + ".No When '' Then '未編號' Else " + wTableName + ".No End As Field1, "
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '開發中' When '1' Then '開發完成(OK)' When '2' Then '開發完成(NG)' Else '開發取消' End As Field2, "

                SQL = SQL + wTableName + ".Division + '--' + "
                SQL = SQL + wTableName + ".Person as Field3, "
                SQL = SQL + wTableName + ".CreateTime as Field4, "
            End If

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
            If pFormNo = "000011" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000012" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "CPSC"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".CPSC   as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "000013" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + "'  ' as Field5, "
                SQL = SQL + "'  ' as Field6, "
                SQL = SQL + "'  ' as Field7, "
                SQL = SQL + "'  ' as Field8, "
            End If
            If pFormNo = "000014" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "OR-No"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".ORNO    as Field8, "
            End If
            If pFormNo = "000015" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "Size-Chain-胴體"
                DataGrid1.Columns.Item(6).HeaderText = "Code"
                DataGrid1.Columns.Item(7).HeaderText = "追加理由"

                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".Spec    as Field6, "
                SQL = SQL + wTableName + ".Code    as Field7, "
                SQL = SQL + wTableName + ".AppendReason    as Field8, "
            End If
            If pFormNo = "000016" Then
                DataGrid1.Columns.Item(4).HeaderText = "色番項目"
                DataGrid1.Columns.Item(5).HeaderText = "購入價/加價"
                DataGrid1.Columns.Item(6).HeaderText = "圖號"
                DataGrid1.Columns.Item(7).HeaderText = "Slider Code"

                SQL = SQL + wTableName + ".ColorItem  as Field5, "
                SQL = SQL + wTableName + ".Pullerprice" + " +'/'+ " + wTableName + ".price as Field6, "
                SQL = SQL + wTableName + ".MapNo   as Field7, "
                SQL = SQL + wTableName + ".SliderCode as Field8, "
            End If
            If pFormNo = "800001" Then
                DataGrid1.Columns.Item(4).HeaderText = "委託書"
                DataGrid1.Columns.Item(5).HeaderText = "狀態"
                DataGrid1.Columns.Item(6).HeaderText = "修改理由類別"
                DataGrid1.Columns.Item(7).HeaderText = "修改理由"

                SQL = SQL + wTableName + ".Sheet   as Field5, "
                SQL = SQL + wTableName + ".Status  as Field6, "
                SQL = SQL + wTableName + ".MReasonType as Field7, "
                SQL = SQL + wTableName + ".MReason as Field8, "
            End If

            If pFormNo = "007001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "圖號"
                DataGrid1.Columns.Item(6).HeaderText = "拉頭品名"
                DataGrid1.Columns.Item(7).HeaderText = "姐妹社"


                SQL = SQL + wTableName + ".Buyer   as Field5, "
                SQL = SQL + wTableName + ".MapNo  as Field6, "
                SQL = SQL + wTableName + ".SliderCode as Field7, "
                SQL = SQL + wTableName + ".YKKGroup as Field8, "
            End If


            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                Dim bc1 As New BoundColumn
                Dim bc2 As New BoundColumn

                bc1.DataField = "Field9"
                bc1.HeaderText = "應收模具費"
                bc2.DataField = "Field10"
                bc2.HeaderText = "模具購入價"

                DataGrid1.Columns.AddAt(8, bc1) ''新增DataGrid欄位
                DataGrid1.Columns.AddAt(9, bc2)
                SetFieldWidth()

                SQL = SQL + wTableName + ".ArMoldFee as Field9, "
                SQL = SQL + wTableName + ".PurMold as Field10, "

                Dim bc3 As New BoundColumn
                bc3.DataField = "Field11"
                bc3.HeaderText = "委託廠商"
                DataGrid1.Columns.AddAt(10, bc3) ''新增DataGrid欄位
                SQL = SQL + wTableName + ".SellVendor as Field11, "

                If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                    Dim bc4 As New BoundColumn
                    bc4.DataField = "Field12"
                    bc4.HeaderText = "外注廠商"
                    DataGrid1.Columns.AddAt(11, bc4) ''新增DataGrid欄位
                    SQL = SQL + wTableName + ".SUPPILER as Field12, "

                End If



            End If
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                SQL = SQL + wTableName + ".MMSSts as MSts, "
            End If

            '20180821 jessica 
            If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
                If DMPSelect.SelectedValue <> "1" Then
                    Dim bc3 As New BoundColumn
                    bc3.DataField = "Field9"
                    bc3.HeaderText = "委託廠商"
                    DataGrid1.Columns.AddAt(8, bc3) ''新增DataGrid欄位
                    SQL = SQL + wTableName + ".SellVendor as Field9, "
                    If pFormNo <> "000011" And pFormNo <> "000012" And pFormNo <> "000015" Then
                        Dim bc4 As New BoundColumn
                        bc4.DataField = "Field10"
                        bc4.HeaderText = "外注廠商"
                        DataGrid1.Columns.AddAt(9, bc4) ''新增DataGrid欄位
                        SQL = SQL + wTableName + ".Suppiler as Field10, "

                    End If


                End If



            End If
            'jessica 20190805 
            'Add-End
            SQL = SQL + " (Select Top 1 Substring(ViewURL,1,CharIndex('?', ViewURL)) From V_WaitHandle_01  Where FormNo = " + wTableName + ".FormNo" + " and FormNo='" + DFormName.SelectedValue + "' ) "
            SQL = SQL + " + 'pFormNo=' + " + wTableName + ".FormNo + '&' + 'pFormSno=' + str(" + wTableName + ".FormSno, Len(" + wTableName + ".FormSno))"
            SQL = SQL + " As ViewURL, "
          
            SQL = SQL + "Case When " + wTableName + ".CreateTime <= '" + SaveTime + "' Then '已封存' Else '未封存' End As WorkFlow, "
            SQL = SQL + " '' As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
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
            SQL = SQL + "Where " + wTableName + ".FormNo <> '" + "@@@@@@@@@@@@@" + "'"
            '表單
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If

            '開發狀態
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If

            '開發完成狀態
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If

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


            '20180821 jessica 增加委託廠商
            'sellvendor
            If DSellVendor.Text <> "委託廠商" And DSellVendor.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SellVendor Like '%" + DSellVendor.Text + "%'"

            End If

            '202401191 jessica 增加外注 
            'sellvendor
            If DSUPPILER.Text <> "外注廠商" And DSUPPILER.Text <> "" Then

                SQL = SQL + " And " + wTableName + ".SUPPILER Like '%" + DSUPPILER.Text + "%'"

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
                        If pFormNo = "000010" Then
                            SQL = SQL + " And " + wTableName + ".SliderGRCode Like '%" + DSliderCode.Text + "%'"
                        Else
                            SQL = SQL + " And " + wTableName + ".SliderCode Like '%" + DSliderCode.Text + "%'"
                        End If
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
            'Code
            If DCode.Text <> "Code" And DCode.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".Code Like '%" + DCode.Text + "%'"
            End If
            ''CompletedTime
            If DMPSelect.SelectedValue = "1" Then   ''判斷是否選取模具費用
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            Else
                SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"
            End If
            'Sort
            SQL = SQL + " Order by " + wTableName + ".No desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            '
            '2012/5/16 Add-Start by Joy
            If pFormNo = "000003" Or pFormNo = "000007" Then
                DBTable1 = DBDataSet2.Tables("WaitHandle")
                If DBTable1.Rows.Count > 0 Then
                    For i = 0 To DBTable1.Rows.Count - 1
                        If DBTable1.Rows(i).Item("MSts") = 1 Then
                            DBTable1.Rows(i).Item("Field1") = DBTable1.Rows(i).Item("Field1") + "-模廢"
                        End If
                    Next
                End If
            End If
            'Add-End
            '
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
        DCpsc.ReadOnly = True
        DCode.ReadOnly = True

        '依賴者
        DPerson.Text = "依賴者"
        'Buyer
        DBuyer.Text = "Buyer"

        'SellVendor 
        DSellVendor.Text = "委託廠商"

        'SellVendor 
        DSUPPILER.Text = "外注廠商"

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
        'Code
        DCode.Text = "Code"

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
            If pFormNo = "000003" Or pFormNo = "000007" Then    ''判斷模具費用是否開啟
                DMPSelect.Enabled = True
            End If
            DBuyer.ReadOnly = False
            DSliderCode.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000010" Then
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
        If pFormNo = "000014" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DCode.ReadOnly = False
        End If
        If pFormNo = "000015" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
        End If
        If pFormNo = "000016" Then
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "800001" Then
        End If
        If pFormNo = "007001" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If

        If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
            DSellVendor.ReadOnly = False
            DSUPPILER.ReadOnly = False
        Else
            DSellVendor.ReadOnly = True
            DSUPPILER.ReadOnly = True
        End If


    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        '
        If DKeepData.SelectedValue = "0" Then
            DataList()      '--未封存資料
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--已封存資料
            Else
                FormDataList()  '--限表單
            End If
        End If
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue

        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DSellVendor.ReadOnly = True
        DSUPPILER.ReadOnly = True
        DMakeMap.ReadOnly = True
        DSpec.ReadOnly = True
        DMapNo.ReadOnly = True
        DSliderCode.ReadOnly = True
        DCpsc.ReadOnly = True
        DCode.ReadOnly = True

        DProgress.Enabled = True    ''預設為啟動
        DProgress.SelectedIndex = 0 ''預設為ALL
        DDivision.Enabled = True    ''預設為啟動
        DSts.SelectedIndex = 0      ''預設為ALL
        DMPSelect.Enabled = False   ''預設模具費用選項為不開啟
        DMPSelect.SelectedIndex = 0 ''預設為一般查詢

        DataGrid1.DataBind()    ''清空DataGrid


        '依賴者
        DPerson.Text = "依賴者"
        'Buyer
        DBuyer.Text = "Buyer"

        'sellvendor
        DSellVendor.Text = "委託廠商"

        'sUPPILER
        DSUPPILER.Text = "外注廠商"


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
        'Code
        DCode.Text = "Code"

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
            If pFormNo = "000003" Or pFormNo = "000007" Then    ''判斷模具費用選項是否要開啟
                DMPSelect.Enabled = True
            End If
            DBuyer.ReadOnly = False
            DSliderCode.ReadOnly = False
            DSpec.ReadOnly = False
            DMapNo.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000010" Then
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
            DCpsc.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
        If pFormNo = "000014" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
            DCode.ReadOnly = False
        End If
        If pFormNo = "000015" Then
            DBuyer.ReadOnly = False
            DSpec.ReadOnly = False
        End If
        If pFormNo = "000016" Then
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If
        If pFormNo = "007001" Then
            DBuyer.ReadOnly = False
            DMapNo.ReadOnly = False
            DSliderCode.ReadOnly = False
        End If


        If pFormNo = "000001" Or pFormNo = "000002" Or pFormNo = "000003" Or pFormNo = "000007" Or pFormNo = "000011" Or pFormNo = "000012" Or pFormNo = "000014" Or pFormNo = "000015" Then
            DSellVendor.ReadOnly = False
            DSUPPILER.ReadOnly = False
        Else
            DSellVendor.ReadOnly = True
            DSUPPILER.ReadOnly = True
        End If

        ''SetFiledDdefault()
    End Sub

    ''*****************************************************************
    ''**
    ''**     檢查各表單篩選欄位狀態
    ''**
    ''*****************************************************************

    ''Sub SetFiledDdefault()
    ''    Dim wTempTextBox As New TextBox
    ''    Dim wDefault As String
    ''    Dim i As Integer

    ''    For i = 1 To 6
    ''        Select Case i
    ''            Case 1
    ''                wTempTextBox = DBuyer
    ''                wDefault = "Buyer"
    ''            Case 2
    ''                wTempTextBox = DMapNo
    ''                wDefault = "圖號"
    ''            Case 3
    ''                wTempTextBox = DMakeMap
    ''                wDefault = "設計者"
    ''            Case 4
    ''                wTempTextBox = DSliderCode
    ''                wDefault = "Slider Code"
    ''            Case 5
    ''                wTempTextBox = DSpec
    ''                wDefault = "Size-Chain-胴體"
    ''            Case 6
    ''                wTempTextBox = DCpsc
    ''                wDefault = "CPSC"
    ''        End Select

    ''        If wTempTextBox.ReadOnly Then       ''若欄位為唯讀,則欄位內容恢復為預設
    ''            wTempTextBox.Text = wDefault
    ''        End If
    ''    Next
    ''End Sub
    ''*****************************************************************
    ''**
    ''**     設定查詢模具費用欄寬
    ''**
    ''*****************************************************************

    Sub SetFieldWidth()     ''設定欄寬
        Dim i As Integer
        Dim TempWidth As Integer
        Dim test As String

        For i = 0 To 10
            Select Case i
                Case 0
                    TempWidth = 90
                Case 1
                    TempWidth = 90
                Case 2
                    TempWidth = 60
                Case 3
                    TempWidth = 30
                Case 4
                    TempWidth = 90
                Case 5
                    TempWidth = 90
                Case 6
                    TempWidth = 90
                Case 7
                    TempWidth = 90
                Case 8
                    TempWidth = 60
                Case 9
                    TempWidth = 60
                Case 10
                    TempWidth = 30
                Case 11
                    TempWidth = 30

            End Select
            DataGrid1.Columns.Item(i).ItemStyle.Width = New Unit(TempWidth)
            DataGrid1.Columns.Item(i).HeaderStyle.Width = New Unit(TempWidth)
        Next
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        '
        '
        If DKeepData.SelectedValue = "0" Then
            DataList()      '--未封存資料
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--已封存資料
            Else
                FormDataList()  '--限表單
            End If
        End If
    End Sub
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        '
        If DKeepData.SelectedValue = "0" Then
            '--未封存資料
            DataList()
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--已封存資料
            Else
                FormDataList()  '--限表單
            End If
        End If

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同
    End Sub

    ''*****************************************************************
    ''**
    ''**     設定欄位狀態
    ''**
    ''*****************************************************************

    Private Sub DMPSelect_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMPSelect.SelectedIndexChanged
        DataGrid1.DataBind()    ''清空DataGrid

        If DMPSelect.SelectedValue = "1" Then   ''判斷模具費用是否選取
            DProgress.SelectedIndex = 2         ''設定為開發完成
            DSts.SelectedIndex = 1              ''設定為OK
            DProgress.Enabled = False
            DDivision.Enabled = False
            DNo.ReadOnly = True
            DPerson.ReadOnly = True
            DBuyer.ReadOnly = True
            DMakeMap.ReadOnly = True
            DSpec.ReadOnly = True
            DMapNo.ReadOnly = True
            DSliderCode.ReadOnly = True
            DCpsc.ReadOnly = True
        Else
            DFormName_SelectedIndexChanged(sender, e)   ''呼叫函數以改變欄位狀態
        End If
    End Sub


    Private Sub DataGrid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
