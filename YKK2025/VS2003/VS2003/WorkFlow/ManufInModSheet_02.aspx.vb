Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ManuInModSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DManuaInSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents DFQAD3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAD2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAC2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAB2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DFQAA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFQAA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFQAA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceJ1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceI1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceG1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceE1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleC2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleB2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQAFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPriceA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSampleA2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSampleFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMoldPoint As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPullerPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPurMold As System.Web.UI.WebControls.TextBox
    Protected WithEvents DArMoldFee As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAuthorizeFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LConfirmFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMaterial As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DManufPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAssembler As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManuaInSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LMapFile As System.Web.UI.WebControls.Image
    Protected WithEvents LQAttachFile2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQAttachFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LSAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Label25 As System.Web.UI.WebControls.Label
    Protected WithEvents Label24 As System.Web.UI.WebControls.Label
    Protected WithEvents Label23 As System.Web.UI.WebControls.Label
    Protected WithEvents Label22 As System.Web.UI.WebControls.Label
    Protected WithEvents Label21 As System.Web.UI.WebControls.Label
    Protected WithEvents Label20 As System.Web.UI.WebControls.Label
    Protected WithEvents Label19 As System.Web.UI.WebControls.Label
    Protected WithEvents Label18 As System.Web.UI.WebControls.Label
    Protected WithEvents Label17 As System.Web.UI.WebControls.Label
    Protected WithEvents Label16 As System.Web.UI.WebControls.Label
    Protected WithEvents Label15 As System.Web.UI.WebControls.Label
    Protected WithEvents Label14 As System.Web.UI.WebControls.Label
    Protected WithEvents DMakeCAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCpsc As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DManufFlow As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCustomerCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLogo As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents Label7 As System.Web.UI.WebControls.Label
    Protected WithEvents Label8 As System.Web.UI.WebControls.Label
    Protected WithEvents Label9 As System.Web.UI.WebControls.Label
    Protected WithEvents Label10 As System.Web.UI.WebControls.Label
    Protected WithEvents Label11 As System.Web.UI.WebControls.Label
    Protected WithEvents Label12 As System.Web.UI.WebControls.Label
    Protected WithEvents Label26 As System.Web.UI.WebControls.Label
    Protected WithEvents Label13 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents DQAA2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAA6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAB6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAC6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQAD6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAB13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAA15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label29 As System.Web.UI.WebControls.Label
    Protected WithEvents DQAB15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAC15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQAD15 As System.Web.UI.WebControls.DropDownList

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
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim NowDateTime As String       '現在日期時間

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "ManufInModSheet_02.aspx"

        SetParameter()          '設定共用參數

        If Not Me.IsPostBack Then   '不是PostBack
            ShowFormData()      '顯示表單資料
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
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
    End Sub


    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufInModSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInModSheet")
        If DBDataSet1.Tables("F_ManufInModSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapFile") <> "" Then          '圖樣
                LMapFile.ImageUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If

            DNo.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Division"))  '擔當
            SetFieldData("Person", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Person"))      '擔當
            SetFieldData("Cpsc", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Cpsc"))          'Cpsc
            SetFieldData("Logo", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Logo"))          'Logo

            DSliderCode.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderCode")       'Slider Code
            DSliderGRCode.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            DSpec.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Spec")                   '規格

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapNo") <> "" Then                 '圖號
                LMapNo.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapNo")
                If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") = "000001" Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormSno"))
                Else
                    LMapNo.NavigateUrl = "MapSheetMod_02.aspx?pFormNo=" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") & _
                                                          "&pFormSno=" & CStr(DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormSno"))
                End If
            Else
                LMapNo.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SAttachFile") <> "" Then       '樣品-其他附件
                LSAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SAttachFile")
            Else
                LSAttachFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile1") <> "" Then      '品質1-其他附件
                LQAttachFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile1")
            Else
                LQAttachFile1.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile2") <> "" Then       '樣品2-其他附件
                LQAttachFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile2")
            Else
                LQAttachFile2.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("RefFile") <> "" Then          '其他附件 
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If

            SetFieldData("Level", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Level"))                '難易度
            DAssembler.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Assembler")                 '組立
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderType1"))    '拉頭種類1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderType2"))    '拉頭種類2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ManufPlace"))      '生產地
            SetFieldData("Material", DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Material"))              '材質
            DBuyer.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SellVendor")   '委託廠商
            DCustomerCode.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("CustomerCode")   'Customer Code

            DMakeCAM.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MakeCAM")   'Make CAM

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ConfirmFile") <> "" Then       '確認書
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '授權書
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If

            DDevReason.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("DevReason")     '開發理由
            DManufFlow.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ManufFlow")     '

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Sample") = 1 Then              '樣品
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOSample "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOSample")
                DBTable1 = DBDataSet1.Tables("ManufCOSample")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Seqno")
                        Case 1
                            DSampleA1.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Spec")   'Spec
                            DSampleA2.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Color")  'Color
                            DSampleA3.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Qty")    'Qty
                        Case 2
                            DSampleB1.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Spec")   'Spec
                            DSampleB2.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Color")  'Color
                            DSampleB3.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Qty")    'Qty
                        Case Else
                            DSampleC1.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Spec")   'Spec
                            DSampleC2.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Color")  'Color
                            DSampleC3.Text = DBDataSet1.Tables("ManufCOSample").Rows(i).Item("Qty")    'Qty
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SampleFile") <> "" Then        '樣品檔
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Price") = 1 Then               '單價
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOPrice "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOPrice")
                DBTable1 = DBDataSet1.Tables("ManufCOPrice")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Seqno")
                        Case 1
                            DPriceA1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceA2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceA3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 2
                            DPriceB1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceB2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceB3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 3
                            DPriceC1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceC2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceC3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 4
                            DPriceD1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceD2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceD3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 5
                            DPriceE1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceE2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceE3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 6
                            DPriceF1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceF2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceF3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 7
                            DPriceG1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceG2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceG3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 8
                            DPriceH1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceH2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceH3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case 9
                            DPriceI1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceI2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceI3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                        Case Else
                            DPriceJ1.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Spec")      'Spec
                            DPriceJ2.SelectedValue = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Currency")  'Currency
                            DPriceJ3.Text = DBDataSet1.Tables("ManufCOPrice").Rows(i).Item("Price")     'Price
                    End Select
                Next
            End If

            DArMoldFee.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ArMoldFee")     '應收模具費
            DPurMold.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("PurMold")         '模具購入費
            DPullerPrice.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("PullerPrice") '引手購入價
            DSuppiler.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Suppiler")       '外注商
            DMoldQty.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MoldQty")         '模型
            DMoldPoint.Text = DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MoldPoint")     '穴取

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Quality1") = 1 Then            '品質1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOQA1 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOQA1")
                DBTable1 = DBDataSet1.Tables("ManufCOQA1")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Seqno")
                        Case 1
                            DQAA1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAA2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAA3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAA4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAA5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAA6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAA7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAA8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAA9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAA10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAA11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAA12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAA13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAA14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                            SetFieldData("DQAA15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
                        Case 2
                            DQAB1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAB2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAB3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAB4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAB5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAB6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAB7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAB8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAB9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAB10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAB11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAB12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAB13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAB14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                            SetFieldData("DQAB15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
                        Case 3
                            DQAC1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAC2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAC3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAC4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAC5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAC6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAC7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAC8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAC9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAC10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAC11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAC12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAC13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAC14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                            SetFieldData("DQAC15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
                        Case Else
                            DQAD1.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Date")
                            DQAD2.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Spec")
                            SetFieldData("DQAD3", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Assembler"))
                            DQAD4.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Surface")
                            DQAD5.Text = DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("GenTani")
                            SetFieldData("DQAD6", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kyoudo"))
                            SetFieldData("DQAD7", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Nyuryoku"))
                            SetFieldData("DQAD8", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Kensin"))
                            SetFieldData("DQAD9", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Water"))
                            SetFieldData("DQAD10", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Dry"))
                            SetFieldData("DQAD11", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Yellow"))
                            SetFieldData("DQAD12", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku1"))
                            SetFieldData("DQAD13", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("Mityaku2"))
                            SetFieldData("DQAD14", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("CPSC"))
                            SetFieldData("DQAD15", DBDataSet1.Tables("ManufCOQA1").Rows(i).Item("EDX"))
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Quality2") = 1 Then            '品質2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_ManufCOQA2 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "ManufCOQA2")
                DBTable1 = DBDataSet1.Tables("ManufCOQA2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Select Case DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Seqno")
                        Case 1
                            DFQAA1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAA2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAA3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                        Case 2
                            DFQAB1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAB2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAB3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                        Case 3
                            DFQAC1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAC2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAC3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                        Case Else
                            DFQAD1.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Date")
                            SetFieldData("DFQAD2", DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("QACheck"))
                            DFQAD3.Text = DBDataSet1.Tables("ManufCOQA2").Rows(i).Item("Remark")
                    End Select
                Next
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAFile") <> "" Then      '測試報告
                LQAFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAFile")
            Else
                LQAFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ContactFile") <> "" Then      '切結書
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If


        End If
        DFormSno.Text = "單號：" & CStr(wFormSno)
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        '部門
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DDivision.Items.Add(ListItem1)
        End If
        '擔當
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DPerson.Items.Add(ListItem1)
        End If
        'Cpsc
        If pFieldName = "Cpsc" Then
            DCpsc.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DCpsc.Items.Add(ListItem1)
        End If
        'Logo
        If pFieldName = "Logo" Then
            DLogo.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DLogo.Items.Add(ListItem1)
        End If

        '難易度
        If pFieldName = "Level" Then
            DLevel.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DLevel.Items.Add(ListItem1)
        End If
        '拉頭種類(內製.外注...)
        If pFieldName = "SliderType1" Then
            DSliderType1.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DSliderType1.Items.Add(ListItem1)
        End If
        '拉頭種類(半成品.成品...)
        If pFieldName = "SliderType2" Then
            DSliderType2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DSliderType2.Items.Add(ListItem1)
        End If
        '生產地
        If pFieldName = "ManufPlace" Then
            DManufPlace.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DManufPlace.Items.Add(ListItem1)
        End If
        '材質
        If pFieldName = "Material" Then
            DMaterial.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DMaterial.Items.Add(ListItem1)
        End If
        'QA--------------------------------------
        'DQAA3
        If pFieldName = "DQAA3" Then
            DQAA3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA3.Items.Add(ListItem1)
        End If
        'DQAA7
        If pFieldName = "DQAA7" Then
            DQAA7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA7.Items.Add(ListItem1)
        End If
        'DQAA8
        If pFieldName = "DQAA8" Then
            DQAA8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA8.Items.Add(ListItem1)
        End If
        'DQAA9
        If pFieldName = "DQAA9" Then
            DQAA9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA9.Items.Add(ListItem1)
        End If
        'DQAA10
        If pFieldName = "DQAA10" Then
            DQAA10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA10.Items.Add(ListItem1)
        End If
        'DQAA11
        If pFieldName = "DQAA11" Then
            DQAA11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA11.Items.Add(ListItem1)
        End If
        'DQAA14
        If pFieldName = "DQAA14" Then
            DQAA14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA14.Items.Add(ListItem1)
        End If
        'DQAA15
        If pFieldName = "DQAA15" Then
            DQAA15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA15.Items.Add(ListItem1)
        End If
        'DQAB3
        If pFieldName = "DQAB3" Then
            DQAB3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB3.Items.Add(ListItem1)
        End If
        'DQAB7
        If pFieldName = "DQAB7" Then
            DQAB7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB7.Items.Add(ListItem1)
        End If
        'DQAB8
        If pFieldName = "DQAB8" Then
            DQAB8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB8.Items.Add(ListItem1)
        End If
        'DQAB9
        If pFieldName = "DQAB9" Then
            DQAB9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB9.Items.Add(ListItem1)
        End If
        'DQAB10
        If pFieldName = "DQAB10" Then
            DQAB10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB10.Items.Add(ListItem1)
        End If
        'DQAB11
        If pFieldName = "DQAB11" Then
            DQAB11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB11.Items.Add(ListItem1)
        End If
        'DQAB14
        If pFieldName = "DQAB14" Then
            DQAB14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB14.Items.Add(ListItem1)
        End If
        'DQAB15
        If pFieldName = "DQAB15" Then
            DQAB15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB15.Items.Add(ListItem1)
        End If
        'DQAC3
        If pFieldName = "DQAC3" Then
            DQAC3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC3.Items.Add(ListItem1)
        End If
        'DQAC7
        If pFieldName = "DQAC7" Then
            DQAC7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC7.Items.Add(ListItem1)
        End If
        'DQAC8
        If pFieldName = "DQAC8" Then
            DQAC8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC8.Items.Add(ListItem1)
        End If
        'DQAC9
        If pFieldName = "DQAC9" Then
            DQAC9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC9.Items.Add(ListItem1)
        End If
        'DQAC10
        If pFieldName = "DQAC10" Then
            DQAC10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC10.Items.Add(ListItem1)
        End If
        'DQAC11
        If pFieldName = "DQAC11" Then
            DQAC11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC11.Items.Add(ListItem1)
        End If
        'DQAC14
        If pFieldName = "DQAC14" Then
            DQAC14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC14.Items.Add(ListItem1)
        End If
        'DQAC15
        If pFieldName = "DQAC15" Then
            DQAC15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC15.Items.Add(ListItem1)
        End If
        'DQAD3
        If pFieldName = "DQAD3" Then
            DQAD3.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD3.Items.Add(ListItem1)
        End If
        'DQAD7
        If pFieldName = "DQAD7" Then
            DQAD7.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD7.Items.Add(ListItem1)
        End If
        'DQAD8
        If pFieldName = "DQAD8" Then
            DQAD8.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD8.Items.Add(ListItem1)
        End If
        'DQAD9
        If pFieldName = "DQAD9" Then
            DQAD9.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD9.Items.Add(ListItem1)
        End If
        'DQAD10
        If pFieldName = "DQAD10" Then
            DQAD10.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD10.Items.Add(ListItem1)
        End If
        'DQAD11
        If pFieldName = "DQAD11" Then
            DQAD11.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD11.Items.Add(ListItem1)
        End If
        'DQAD14
        If pFieldName = "DQAD14" Then
            DQAD14.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD14.Items.Add(ListItem1)
        End If
        'DQAD15
        If pFieldName = "DQAD15" Then
            DQAD15.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD15.Items.Add(ListItem1)
        End If
        '------------------
        'DQAA6
        If pFieldName = "DQAA6" Then
            DQAA6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA6.Items.Add(ListItem1)
        End If
        'DQAA12
        If pFieldName = "DQAA12" Then
            DQAA12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA12.Items.Add(ListItem1)
        End If
        'DQAA13
        If pFieldName = "DQAA13" Then
            DQAA13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAA13.Items.Add(ListItem1)
        End If

        'DQAB6
        If pFieldName = "DQAB6" Then
            DQAB6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB6.Items.Add(ListItem1)
        End If
        'DQAB12
        If pFieldName = "DQAB12" Then
            DQAB12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB12.Items.Add(ListItem1)
        End If
        'DQAB13
        If pFieldName = "DQAB13" Then
            DQAB13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAB13.Items.Add(ListItem1)
        End If

        'DQAC6
        If pFieldName = "DQAC6" Then
            DQAC6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC6.Items.Add(ListItem1)
        End If
        'DQAC12
        If pFieldName = "DQAC12" Then
            DQAC12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC12.Items.Add(ListItem1)
        End If
        'DQAC13
        If pFieldName = "DQAC13" Then
            DQAC13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAC13.Items.Add(ListItem1)
        End If

        'DQAD6
        If pFieldName = "DQAD6" Then
            DQAD6.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD6.Items.Add(ListItem1)
        End If
        'DQAD12
        If pFieldName = "DQAD12" Then
            DQAD12.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD12.Items.Add(ListItem1)
        End If
        'DQAD13
        If pFieldName = "DQAD13" Then
            DQAD13.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQAD13.Items.Add(ListItem1)
        End If
        '----------------------------------

        'QA-1 
        'DFQAA2
        If pFieldName = "DFQAA2" Then
            DFQAA2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAA2.Items.Add(ListItem1)
        End If
        'DFQAB2
        If pFieldName = "DFQAB2" Then
            DFQAB2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAB2.Items.Add(ListItem1)
        End If
        'DFQAC2
        If pFieldName = "DFQAC2" Then
            DFQAC2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAC2.Items.Add(ListItem1)
        End If
        'DFQAD2
        If pFieldName = "DFQAD2" Then
            DFQAD2.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DFQAD2.Items.Add(ListItem1)
        End If

    End Sub

End Class
