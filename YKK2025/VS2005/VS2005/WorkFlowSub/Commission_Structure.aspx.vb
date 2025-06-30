Imports System.Data
Imports System.Data.OleDb

Partial Class Commission_Structure
    Inherits System.Web.UI.Page

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
    Dim NowDateTime As String       '現在日期時間
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cookies("PGM").Value = "Commission_Structure.aspx"
        TreeView1.Visible = False
        DMsgBox.Text = ""

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetLastUpdateTime()
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
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetLastUpdateTime()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        DLastUpdateTime.Text = ""
        SQL = "SELECT CreateTime From W_CommissionStructure "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Sheet")
        If DBDataSet1.Tables("Sheet").Rows.Count > 0 Then
            DLastUpdateTime.Text = DBDataSet1.Tables("Sheet").Rows(0).Item("CreateTime") + "　時點"
        End If

        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '依賴部門
        OleDbConnection1.Open()
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
        OleDbConnection1.Close()

        Search_Item_Attribute()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定搜尋欄位屬性
    '**
    '*****************************************************************
    Sub Search_Item_Attribute()
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
        DBuyer.ReadOnly = False
        DMakeMap.ReadOnly = False
        DSliderCode.ReadOnly = False
        DSpec.ReadOnly = False
        DMapNo.ReadOnly = False
        DCpsc.ReadOnly = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Go
    '**
    '*****************************************************************
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        Dim ErrCode As Integer = SearchItem_Check()
        DMsgBox.Text = ""
        TreeView1.Visible = False

        If ErrCode = 0 Then
            Session.Clear()
            TreeView1.Nodes.Clear()
            InitTree()     '初始化Tree   
            BuildTree()    '建立Tree內容  
        Else
            DMsgBox.Text = "-->  篩選項目中，需指定１個以上的篩選項目"
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     TreeView初期設定
    '**
    '*****************************************************************
    Sub InitTree()
        Dim tmpNote As New TreeNode  '定義一個TreeNode並實體化
        tmpNote.Text = "委託書"  ''設定【根目錄】相關屬性內容
        tmpNote.Value = "0"
        tmpNote.NavigateUrl = ""
        tmpNote.Target = ""

        TreeView1.Nodes.Add(tmpNote)  'Tree建立該Node
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     TreeView展開根目錄節點
    '**
    '*****************************************************************
    Sub BuildTree()
        Dim RootNode As TreeNode  '取得根目錄節點   
        RootNode = TreeView1.Nodes(0)

        Dim ErrCode As Integer = GetDataTable()
        If ErrCode = 0 Then
            TreeView1.Visible = True
            Dim rc As String
            rc = AddNodes(RootNode, 0)  '呼叫建立子節點的函數
        Else
            If ErrCode = 1 Then DMsgBox.Text = "-->  無符合的資料，請檢視搜尋條件"
            If ErrCode = 2 Then DMsgBox.Text = "-->  符合的資料過多，請檢視搜尋條件"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     TreeView展開子節點
    '**
    '*****************************************************************
    Function AddNodes(ByRef tNode As TreeNode, ByVal PId As Integer) As String
        Dim wHttp As String = System.Configuration.ConfigurationManager.AppSettings("Http")  '網站
        Dim Sheet As DataTable  '定義DataTable  

        Sheet = Session("DBTable")  '從Session中取得DataTable

        Dim rows() As DataRow  '定義DataRow承接DataTable篩選的結果
        Dim filterExpr As String  '定義篩選的條件
        filterExpr = "ParnetId = " & PId
        rows = Sheet.Select(filterExpr)  '資料篩選並把結果傳入Rows 

        If rows.GetUpperBound(0) >= 0 Then  '如果篩選結果有資料
            Dim row As DataRow
            Dim tmpNodeId As Long
            Dim tmpsText As String
            Dim tmpsValue As String
            Dim tmpsUrl As String
            Dim tmpsTarget As String
            Dim NewNode As TreeNode
            Dim rc As String
            For Each row In rows   '逐筆取出篩選後資料
                tmpNodeId = row(0)   '放入相關變數中   
                tmpsText = row(2)
                tmpsValue = row(3)
                If row(4) <> "" Then
                    tmpsUrl = wHttp & row(4)
                Else
                    tmpsUrl = ""
                End If

                tmpsTarget = row(5)

                NewNode = New TreeNode   '實體化新節點 
                NewNode.Text = tmpsText   '設定節點各屬性
                NewNode.Value = tmpsValue
                NewNode.NavigateUrl = tmpsUrl
                NewNode.Target = tmpsTarget

                tNode.ChildNodes.Add(NewNode)   '將節點加入Tree中

                rc = AddNodes(NewNode, tmpNodeId)  '呼叫遞回取得子節點 
            Next
        End If
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     TreeView用DB資料
    '**
    '*****************************************************************
    Function GetDataTable() As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "SELECT NoteId, ParnetId, sText, sValue, sURL, sTarget FROM W_CommissionStructure "
        SQL = SQL & "Where Sts < '10' "
        '部門
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
        End If
        '依賴人
        If DPerson.Text <> "依賴者" And DPerson.Text <> "" Then
            SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
        End If
        'Buyer
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
            SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
        End If
        '設計者
        If DMakeMap.Text <> "設計者" And DMakeMap.Text <> "" Then
            SQL = SQL + " And MakeMap Like '%" + DMakeMap.Text + "%'"
        End If
        'Spec
        If DSpec.Text <> "Size-Chain-胴體" And DSpec.Text <> "" Then
            SQL = SQL + " And Spec Like '%" + DSpec.Text + "%'"
        End If
        'MapNo
        If DMapNo.Text <> "圖號" And DMapNo.Text <> "" Then
            SQL = SQL + " And MapNo Like '%" + DMapNo.Text + "%'"
        End If
        'Cpsc
        If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then
            SQL = SQL + " And Cpsc Like '%" + DCpsc.Text + "%'"
        End If
        'No
        If DNo.Text <> "委託單No." And DNo.Text <> "" Then
            SQL = SQL + " And No Like '%" + DNo.Text + "%'"
        End If
        'Slider Code
        If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then
            SQL = SQL + " And SliderCode Like '%" + DSliderCode.Text + "%'"
        End If

        SQL = SQL & "Order by NoteId, ParnetId, sText, sValue, sURL, sTarget "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Sheet")
        DBTable1 = DBDataSet1.Tables("Sheet")
        If DBTable1.Rows.Count > 0 Then
            If DBTable1.Rows.Count > 100000 Then
                GetDataTable = 2
            Else
                Session("DBTable") = DBTable1
                GetDataTable = 0
            End If
        Else
            GetDataTable = 1
        End If

        OleDbConnection1.Close()

    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Check Search Item 
    '**
    '*****************************************************************
    Function SearchItem_Check() As Integer
        Dim HaveItem As Boolean = False
        SearchItem_Check = 1

        '部門
        If DDivision.SelectedValue <> "ALL" Then HaveItem = True
        '依賴人
        If DPerson.Text <> "依賴者" And DPerson.Text <> "" Then HaveItem = True
        'Buyer
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then HaveItem = True
        '設計者
        If DMakeMap.Text <> "設計者" And DMakeMap.Text <> "" Then HaveItem = True
        'Spec
        If DSpec.Text <> "Size-Chain-胴體" And DSpec.Text <> "" Then HaveItem = True
        'MapNo
        If DMapNo.Text <> "圖號" And DMapNo.Text <> "" Then HaveItem = True
        'Cpsc
        If DCpsc.Text <> "CPSC" And DCpsc.Text <> "" Then HaveItem = True
        'No
        If DNo.Text <> "委託單No." And DNo.Text <> "" Then HaveItem = True
        'Slider Code
        If DSliderCode.Text <> "Slider Code" And DSliderCode.Text <> "" Then HaveItem = True

        If HaveItem Then SearchItem_Check = 0
    End Function

End Class
