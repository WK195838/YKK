Imports System.Data
Imports System.Data.OleDb

Public Class DevNoPicker
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BKey As System.Web.UI.WebControls.Button
    Protected WithEvents DKey As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid

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
            If Request.Cookies("DevNo").Value <> "" Then
                DKey.Text = Request.Cookies("DevNo").Value
                pKey = Request.Cookies("DevNo").Value
            Else
                DKey.Text = ""
                pKey = "ALL"
            End If
            MapData()  '取得資料
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandSource.CommandName = "Select" Then  '點選選取檢查
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '所選取的Map No
            Dim wFormNo As String = ""
            Dim wFormSno As String = ""
            Dim wLevel As String = ""
            Dim wDivision As String = ""
            Dim wPerson As String = ""
            Dim wSliderGRCode As String = ""
            Dim wSliderCode As String = ""
            Dim wSuppiler As String = ""
            Dim wMakeCAM As String = ""
            Dim wMapNo As String = ""
            Dim wMapNo1 As String = ""              '外注色番追加表使用
            Dim wBuyer As String = ""
            Dim wSellVendor As String = ""

            Dim SQL As String
            Dim DBDataSet1 As New DataSet

            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()

            If Request.QueryString("field") = "In" Then
                SQL = "Select FormNo, FormSno, Level, Division, Person, SliderCode, SliderGRCode, Suppiler, MakeCAM, MapNo, Buyer, SellVendor From F_ManufInSheet "
            End If
            If Request.QueryString("field") = "Out" Then
                SQL = "Select FormNo, FormSno, Level, Division, Person, SliderCode, SliderGRCode, Suppiler, MapNo, MapNo1, Buyer, SellVendor From F_ManufOutSheet "
            End If
            If Request.QueryString("field") = "Import" Then
                SQL = "Select FormNo, FormSno, Division, Person, SliderCode, SliderGRCode, Buyer, SellVendor From F_ImportSheet "
            End If
            SQL = SQL & " Where Sts =  '1' "
            If pKey <> "ALL" Then
                SQL = SQL & "   And  No = '" & Key & "'"
                'SQL = SQL & "   And  No Like '%" & Key & "%'"
            End If

            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_ManufSheet")
            If DBDataSet1.Tables("F_ManufSheet").Rows.Count > 0 Then
                wFormNo = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("FormNo")
                wFormSno = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("FormSno")
                wSliderGRCode = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("SliderGRCode")
                wSliderCode = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("SliderCode")
                wBuyer = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("Buyer")
                wSellVendor = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("SellVendor")

                If Request.QueryString("field") <> "Import" Then
                    If Request.QueryString("field") = "In" Then
                        wMakeCAM = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("MakeCAM")
                    Else
                        wMapNo1 = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("MapNo1")
                    End If
                    wLevel = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("Level")
                    wSuppiler = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("Suppiler")
                    wMapNo = DBDataSet1.Tables("F_ManufSheet").Rows(0).Item("MapNo")
                End If
            End If
            'DB連結關閉
            OleDbConnection1.Close()

            '所選取的Map No回父視窗並關閉
            Dim Cmd As String
            Response.Cookies("DevNo").Value = Key

            If Request.QueryString("pFormNo") = "000004" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.opener.document.{12}.value = '{13:d}'; window.opener.document.{14}.value = '{15:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSuppiler", wSuppiler, "Form1.DSliderCode", wSliderCode, "Form1.DMakeCAM", wMakeCAM, "Form1.DMapNo", wMapNo, "Form1.DLevel", wLevel)
            End If

            If Request.QueryString("pFormNo") = "000005" Or Request.QueryString("pFormNo") = "000009" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSliderCode", wSliderCode, "Form1.DMapNo", wMapNo, "Form1.DLevel", wLevel)
            End If


            If Request.QueryString("pFormNo") = "000008" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.opener.document.{12}.value = '{13:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSuppiler", wSuppiler, "Form1.DSliderCode", wSliderCode, "Form1.DMapNo", wMapNo, "Form1.DLevel", wLevel)
            End If

            If Request.QueryString("pFormNo") = "000010" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSliderGRCode", wSliderGRCode)
            End If

            If Request.QueryString("pFormNo") = "000012" Then
                If Request.QueryString("field") <> "Import" Then
                    Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.opener.document.{12}.value = '{13:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DBuyer", wBuyer, "Form1.DSellVendor", wSellVendor, "Form1.DSliderCode", wSliderCode, "Form1.DMapNo", wMapNo)
                Else
                    Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DBuyer", wBuyer, "Form1.DSellVendor", wSellVendor, "Form1.DSliderCode", wSliderCode)
                End If
            End If

            If Request.QueryString("pFormNo") = "000013" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSliderCode", wSliderCode)
            End If

            ''表面處理委託書
            If Request.QueryString("pFormNo") = "000014" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno)
                'Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2} = '{3:d}'; window.opener.document.{4} = '{5:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.wOFormNoTmp", wFormNo, "Form1.wOFormSnoTmp", wFormSno)
            End If

            ''外注色番追加表
            If Request.QueryString("pFormNo") = "000016" Then
                If Request.QueryString("field") <> "Import" Then
                    If Request.QueryString("field") = "Out" And wMapNo1 > wMapNo Then wMapNo = wMapNo1
                    Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSliderCode", wSliderCode, "Form1.DMapNo", wMapNo, "Form1.DLevel", wLevel)
                Else
                    Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}';window.close();</script>", "Form1.DNo", Key, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DSliderCode", wSliderCode, "Form1.DLevel", "0")
                End If
            End If

            Response.Write(Cmd)
        End If
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        'DataGrid1.DataBind()
        MapData()
    End Sub

    Sub MapData()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim SQL As String




        pKey = DKey.Text

        SQL = "SELECT "   '取得DB資料
        SQL = SQL & "No, SliderGRCode, "
        SQL = SQL & "FormNo +  '-' + RTrim(LTrim(str(formsno))) as FormNoDesc, "

        If Request.QueryString("field") = "In" Then
            SQL = SQL & "'ManufInSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_ManufInSheet "
        End If
        If Request.QueryString("field") = "Out" Then
            SQL = SQL & "'ManufOutSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_ManufOutSheet "
        End If
        If Request.QueryString("field") = "Import" Then
            SQL = SQL & "'ImportSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_ImportSheet "
        End If
        SQL = SQL & " Where Sts =  '1' "

        If Request.QueryString("pFormNo") = "000004" Or Request.QueryString("pFormNo") = "000008" Then
            SQL = SQL + "And   OPSts = '0' "
        End If
        If Request.QueryString("pFormNo") = "000005" Or Request.QueryString("pFormNo") = "000009" Or Request.QueryString("pFormNo") = "000013" Then
            SQL = SQL + "And   CTSts = '0' "
        End If
        If Request.QueryString("pFormNo") = "000010" Then
            SQL = SQL + "And   SDSts = '0' "
        End If
        ' Jessica  SPD -表面處理委託書 同時追加多張問題 2013/11/27
        ' If Request.QueryString("pFormNo") = "000014" Then
        ' SQL = SQL + "And   SFSts = '0' "
        ' End If

        ' Jessica  SPD -外注色番追加表 同時追加多張問題 2012/02/06
        '   If Request.QueryString("pFormNo") = "000016" Then
        '  SQL = SQL + "And   CASts = '0' "
        '  End If

        If (pKey <> "ALL") And (pKey <> "") Then
            SQL = SQL + "And No Like '%" & pKey & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufSheet")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub BKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BKey.Click
        DataGrid1.CurrentPageIndex = 0  'DataGrid跳上下頁
        pKey = DKey.Text
        MapData()
    End Sub
End Class
