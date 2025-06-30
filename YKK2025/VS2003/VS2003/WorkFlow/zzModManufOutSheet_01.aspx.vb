Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ModManufOutSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DManuaInSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSave As System.Web.UI.WebControls.Button
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents LMapFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LMapFile2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAssembler As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderType1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSliderType2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DManufPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterialDetail As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterial As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterialOther As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents LConfirmFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LAuthorizeFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSliderCode As System.Web.UI.WebControls.Button
    Protected WithEvents BSample As System.Web.UI.WebControls.Button
    Protected WithEvents BPrice As System.Web.UI.WebControls.Button
    Protected WithEvents DArMoldFee As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPurMold As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPullerPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMoldPoint As System.Web.UI.WebControls.TextBox
    Protected WithEvents BQuality1 As System.Web.UI.WebControls.Button
    Protected WithEvents BQuality2 As System.Web.UI.WebControls.Button
    Protected WithEvents LQAAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LSampleFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BMMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQuality1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQuality2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSliderCode As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LSample As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LPrice As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQuality1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQuality2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents wMapFile1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents wMapFile2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents wConfirmFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents wAuthorizeFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents wSampleFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents wQAAttachFile As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManuaInSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQAAttachFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DAuthorizeFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DConfirmFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DMapFile2 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DMapFile1 As System.Web.UI.HtmlControls.HtmlInputFile

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
    Dim FieldName(44) As String     '各欄位
    Dim Attribute(44) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wOFormNo As String          '原表單號碼
    Dim wOFormSno As Integer        '原表單流水號
    Dim rFormNo As String           '傳回表單號碼
    Dim rFormSno As Integer         '傳回表單流水號
    Dim wStep As Integer            '工程代碼
    Dim NowDateTime As String       '現在日期時間

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示
            If wFormSno > 0 Then    '判斷是否有資料
                ShowFormData()      '顯示表單資料
            Else
                ShowOriFormData()   '顯示原表單資料
            End If
            SetPopupFunction()      '設定彈出視窗事件
        Else
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查
            'ShowMessage()           '上傳資料檢查及顯示訊息
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
                      CStr(DateTime.Now.Second)       '現在日時
        wFormNo = Request.QueryString("pFormNo")
        wFormSno = Request.QueryString("pFormSno")
        wOFormNo = Request.QueryString("pOFormNo")    '原表單號碼
        wOFormSno = Request.QueryString("pOFormSno")  '原表單流水號
        wStep = Request.QueryString("pStep")          '工程代碼
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        'Check附圖-1
        If DMapFile1.Visible Then
            If DMapFile1.PostedFile.FileName <> "" Then
                Message = "附圖-1"
            End If
        End If
        'Check附圖-2
        If DMapFile2.Visible Then
            If DMapFile2.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "附圖-2"
                Else
                    Message = Message & ", " & "附圖-2"
                End If
            End If
        End If
        'Check確認書
        If DConfirmFile.Visible Then
            If DConfirmFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "確認書"
                Else
                    Message = Message & ", " & "確認書"
                End If
            End If
        End If
        'Check授權書
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "授權書"
                Else
                    Message = Message & ", " & "授權書"
                End If
            End If
        End If
        'Check樣品圖
        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "樣品圖"
                Else
                    Message = Message & ", " & "樣品圖"
                End If
            End If
        End If
        'CheckQA附件
        If DQAAttachFile.Visible Then
            If DQAAttachFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "QA附件"
                Else
                    Message = Message & ", " & "QA附件"
                End If
            End If
        End If

        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 1120
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        If wFormSno > 0 And wStep > 1 Then    '判斷是否[簽核]
            'Sheet顯示
            DManuaInSheet1.Visible = True   '表單Sheet-1
            DManuaInSheet2.Visible = True   '表單Sheet-2
            '按鈕位置
            BSave.Style.Add("Top", Top)    '儲存按鈕
            DFormSno.Style.Add("Top", Top) '單號
        Else
            '按鈕位置
            BSave.Style.Add("Top", Top)    '儲存按鈕
            DFormSno.Style.Add("Top", Top) '單號
        End If

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '900003' "
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            '電子簽章未使用
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Attach") = 1 Then
                'DRefMapFile.Visible = True
                'DMapFile.Visible = True
            Else
                'DRefMapFile.Visible = False
                'DMapFile.Visible = False
            End If
            '儲存按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveFun") = 1 Then
                BSave.Visible = True
            Else
                BSave.Visible = False
            End If
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String
        If wOFormNo = "000003" Then
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Else
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufOutFilePath")
        End If

        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ModManufSheet "
        SQL = SQL & " Where Sts = 0 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
        If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Person"))      '擔當
            DSliderGRCode.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderCode") = 1 Then              'Slider Code
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SliderList "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SliderList")
                DBTable1 = DBDataSet1.Tables("SliderList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("SliderCode") & "]"
                    If DSliderCode.Text = "" Then
                        DSliderCode.Text = Str
                    Else
                        DSliderCode.Text = DSliderCode.Text + ";" + Str
                    End If
                Next
                LSliderCode.NavigateUrl = "SliderCodeList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LSliderCode.Visible = False
            End If

            DSpec.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Spec")               '規格(Size,ChainType,胴體的集合)
            DMapNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapNo")             '圖號
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1") <> "" Then          '圖檔1
                LMapFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1")
            Else
                LMapFile1.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2") <> "" Then          '圖檔2 
                LMapFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2")
            Else
                LMapFile2.Visible = False
            End If
            SetFieldData("Level", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Level"))                '難易度
            DAssembler.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Assembler")                 '組立
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType1"))    '拉頭種類1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType2"))    '拉頭種類2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ManufPlace"))      '生產地
            SetFieldData("Material", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Material"))              '材質
            SetFieldData("MaterialDetail", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialDetail"))  '材質細項
            DMaterialOther.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialOther")             '材質其他
            DBuyer.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SellVendor")   '委託廠商
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile") <> "" Then       '確認書
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '授權書
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If
            DDevReason.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("DevReason")     '開發理由

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Sample") = 1 Then              '樣品
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SampleList "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SampleList")
                DBTable1 = DBDataSet1.Tables("SampleList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Class") & "," & DBTable1.Rows(i).Item("Qty") & "]"
                    If DSample.Text = "" Then
                        DSample.Text = Str
                    Else
                        DSample.Text = DSample.Text + ";" + Str
                    End If
                Next
                LSample.NavigateUrl = "SampleList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LSample.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile") <> "" Then        '樣品檔
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Price") = 1 Then               '單價
                Dim DBTable1 As DataTable
                SQL = "Select * From F_PriceList "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "PriceList")
                DBTable1 = DBDataSet1.Tables("PriceList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Price") & "]"
                    If DPrice.Text = "" Then
                        DPrice.Text = Str
                    Else
                        DPrice.Text = DPrice.Text + ";" + Str
                    End If
                Next
                LPrice.NavigateUrl = "PriceList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LPrice.Visible = False
            End If

            DArMoldFee.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ArMoldFee")     '應收模具費
            DPurMold.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PurMold")         '模具購入費
            DPullerPrice.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PullerPrice") '引手購入價
            DMoldName.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldName")       '模具名稱
            DMoldQty.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty")         '模型
            DMoldPoint.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldPoint")     '穴取

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality1") = 1 Then            '品質1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList1 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList1")
                DBTable1 = DBDataSet1.Tables("QAList1")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & _
                            DBTable1.Rows(i).Item("Date") & "," & _
                            DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                            DBTable1.Rows(i).Item("Assembler") & "," & DBTable1.Rows(i).Item("Surface1") & "," & DBTable1.Rows(i).Item("Surface2") & "," & DBTable1.Rows(i).Item("Gentani") & "," & _
                            DBTable1.Rows(i).Item("Kyoudo") & "," & DBTable1.Rows(i).Item("Nyuryoku") & "," & DBTable1.Rows(i).Item("Kensin") & "," & _
                            DBTable1.Rows(i).Item("Water") & "," & DBTable1.Rows(i).Item("Dry") & "," & _
                            DBTable1.Rows(i).Item("Yellow") & "," & DBTable1.Rows(i).Item("Mityaku") & "," & DBTable1.Rows(i).Item("CPSC") & _
                          "]"
                    If DQuality1.Text = "" Then
                        DQuality1.Text = Str
                    Else
                        DQuality1.Text = DQuality1.Text + ";" + Str
                    End If
                Next
                LQuality1.NavigateUrl = "Quality1List.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LQuality1.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality2") = 1 Then            '品質2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList2 "
                SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList2")
                DBTable1 = DBDataSet1.Tables("QAList2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Date") & "," & DBTable1.Rows(i).Item("QACheck") & "," & DBTable1.Rows(i).Item("Remark") & "]"
                    If DQuality2.Text = "" Then
                        DQuality2.Text = Str
                    Else
                        DQuality2.Text = DQuality2.Text + ";" + Str
                    End If
                Next
                LQuality2.NavigateUrl = "Quality2List.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno
            Else
                LQuality2.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile") <> "" Then      '品質附件
                LQAAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile")
            Else
                LQAAttachFile.Visible = False
            End If

            DFormSno.Text = "單號：" & CStr(wFormSno) '單號
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowOriFormData()
        Dim Path As String
        If wOFormNo = "000003" Then
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInFilePath")
        Else
            Path = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufOutFilePath")
        End If
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufOutSheet "
        SQL = SQL & " Where Sts = 2 "
        SQL = SQL & "   And FormNo =  '" & wOFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
        If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Person"))      '擔當
            DSliderGRCode.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderGRCode")   'Slider Group Code

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderCode") = 1 Then              'Slider Code
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SliderList "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SliderList")
                DBTable1 = DBDataSet1.Tables("SliderList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("SliderCode") & "]"
                    If DSliderCode.Text = "" Then
                        DSliderCode.Text = Str
                    Else
                        DSliderCode.Text = DSliderCode.Text + ";" + Str
                    End If
                Next
                LSliderCode.NavigateUrl = "SliderCodeList.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LSliderCode.Visible = False
            End If

            DSpec.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Spec")               '規格(Size,ChainType,胴體的集合)
            DMapNo.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapNo")             '圖號
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1") <> "" Then          '圖檔1
                LMapFile1.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1")
                wMapFile1.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1")
            Else
                LMapFile1.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2") <> "" Then          '圖檔2 
                LMapFile2.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2")
                wMapFile2.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2")
            Else
                LMapFile2.Visible = False
            End If
            SetFieldData("Level", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Level"))                '難易度
            DAssembler.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Assembler")                 '組立
            SetFieldData("SliderType1", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType1"))    '拉頭種類1
            SetFieldData("SliderType2", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType2"))    '拉頭種類2
            SetFieldData("ManufPlace", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ManufPlace"))      '生產地
            SetFieldData("Material", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Material"))              '材質
            SetFieldData("MaterialDetail", DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialDetail"))  '材質細項
            DMaterialOther.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialOther")             '材質其他
            DBuyer.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Buyer")             'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SellVendor")   '委託廠商
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile") <> "" Then       '確認書
                LConfirmFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile")
                wConfirmFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile")
            Else
                LConfirmFile.Visible = False
            End If
            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile") <> "" Then     '授權書
                LAuthorizeFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile")
                wAuthorizeFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile")
            Else
                LAuthorizeFile.Visible = False
            End If
            DDevReason.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("DevReason")     '開發理由

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Sample") = 1 Then              '樣品
                Dim DBTable1 As DataTable
                SQL = "Select * From F_SampleList "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "SampleList")
                DBTable1 = DBDataSet1.Tables("SampleList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Class") & "," & DBTable1.Rows(i).Item("Qty") & "]"
                    If DSample.Text = "" Then
                        DSample.Text = Str
                    Else
                        DSample.Text = DSample.Text + ";" + Str
                    End If
                Next
                LSample.NavigateUrl = "SampleList.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LSample.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile") <> "" Then        '樣品檔
                LSampleFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile")
                wSampleFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile")
            Else
                LSampleFile.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Price") = 1 Then               '單價
                Dim DBTable1 As DataTable
                SQL = "Select * From F_PriceList "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "PriceList")
                DBTable1 = DBDataSet1.Tables("PriceList")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                                DBTable1.Rows(i).Item("Price") & "]"
                    If DPrice.Text = "" Then
                        DPrice.Text = Str
                    Else
                        DPrice.Text = DPrice.Text + ";" + Str
                    End If
                Next
                LPrice.NavigateUrl = "PriceList.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LPrice.Visible = False
            End If

            DArMoldFee.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ArMoldFee")     '應收模具費
            DPurMold.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PurMold")         '模具購入費
            DPullerPrice.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PullerPrice") '引手購入價
            DMoldName.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldName")       '模具名稱
            DMoldQty.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty")         '模型
            DMoldPoint.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldPoint")     '穴取

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality1") = 1 Then            '品質1
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList1 "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList1")
                DBTable1 = DBDataSet1.Tables("QAList1")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & _
                            DBTable1.Rows(i).Item("Date") & "," & _
                            DBTable1.Rows(i).Item("Size") & "," & DBTable1.Rows(i).Item("ChainType") & "," & DBTable1.Rows(i).Item("Class") & "," & _
                            DBTable1.Rows(i).Item("Assembler") & "," & DBTable1.Rows(i).Item("Surface1") & "," & DBTable1.Rows(i).Item("Surface2") & "," & DBTable1.Rows(i).Item("Gentani") & "," & _
                            DBTable1.Rows(i).Item("Kyoudo") & "," & DBTable1.Rows(i).Item("Nyuryoku") & "," & DBTable1.Rows(i).Item("Kensin") & "," & _
                            DBTable1.Rows(i).Item("Water") & "," & DBTable1.Rows(i).Item("Dry") & "," & _
                            DBTable1.Rows(i).Item("Yellow") & "," & DBTable1.Rows(i).Item("Mityaku") & "," & DBTable1.Rows(i).Item("CPSC") & _
                          "]"
                    If DQuality1.Text = "" Then
                        DQuality1.Text = Str
                    Else
                        DQuality1.Text = DQuality1.Text + ";" + Str
                    End If
                Next
                LQuality1.NavigateUrl = "Quality1List.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LQuality1.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality2") = 1 Then            '品質2
                Dim DBTable1 As DataTable
                SQL = "Select * From F_QAList2 "
                SQL = SQL & " Where FormNo =  '" & wOFormNo & "'"
                SQL = SQL & "   And FormSno =  '" & CStr(wOFormSno) & "'"
                SQL = SQL & " Order by Seqno"
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "QAList2")
                DBTable1 = DBDataSet1.Tables("QAList2")
                For i = 0 To DBTable1.Rows.Count - 1
                    Str = "[" & DBTable1.Rows(i).Item("Date") & "," & DBTable1.Rows(i).Item("QACheck") & "," & DBTable1.Rows(i).Item("Remark") & "]"
                    If DQuality2.Text = "" Then
                        DQuality2.Text = Str
                    Else
                        DQuality2.Text = DQuality2.Text + ";" + Str
                    End If
                Next
                LQuality2.NavigateUrl = "Quality2List.aspx?pFormNo=" & wOFormNo & "&pFormSno=" & wOFormSno
            Else
                LQuality2.Visible = False
            End If

            If DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile") <> "" Then      '品質附件
                LQAAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile")
                wQAAttachFile.Text = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile")
            Else
                LQAAttachFile.Visible = False
            End If

            DFormSno.Text = "單號：" & CStr(wOFormSno)       '單號
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"  '日期選擇
        BSliderCode.Attributes("onclick") = "SliderPicker('Form1.DSliderCode');"  'Slider Code
        BSpec.Attributes("onclick") = "SpecPicker('Form1.DSpec');"      '規格
        BOMapNo.Attributes("onclick") = "MapPicker('Ori');"    '原始圖號
        BMMapNo.Attributes("onclick") = "MapPicker('Mod');"    '修改圖號
        BSample.Attributes("onclick") = "SamplePicker('Form1.DSample');"    '樣品
        BPrice.Attributes("onclick") = "PricePicker('Form1.DPrice');"    '單價
        BQuality1.Attributes("onclick") = "QA1Picker('Form1.DQuality1');"    '品質分析1
        BQuality2.Attributes("onclick") = "QA2Picker('Form1.DQuality2');"    '品質分析1
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        Dim oFieldAttribute As Object
        oFieldAttribute = Server.CreateObject("GetFieldAttb.WFField")
        oFieldAttribute.pFormNo = "900003"     '表單號碼
        oFieldAttribute.pStep = wStep          '工程關卡號碼
        oFieldAttribute.GetFieldAttribute(FieldName, Attribute)       '欄位名,欄位屬性

        SetFieldAttribute(pPost)     '表單各欄位屬性及欄位輸入檢查等設定
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.PeachPuff
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '修改+檢查
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "異常：需輸入Ｎｏ")
                DNo.Visible = True
            Case 2  '修改
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""
        '日期
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.PeachPuff
                DDate.ReadOnly = True
                DDate.Visible = True
                BDate.Visible = False
            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DDate.Visible = True
                BDate.Visible = True
            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.Visible = True
                BDate.Visible = True
            Case Else   '隱藏
                DDate.Visible = False
                BDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)
        '部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.PeachPuff
                DDivision.Enabled = False
                DDivision.Visible = True
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")
        '擔當
        Select Case FindFieldInf("Person")
            Case 0  '顯示
                DPerson.BackColor = Color.PeachPuff
                DPerson.Enabled = False
                DPerson.Visible = True
            Case 1  '修改+檢查
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "異常：需輸入擔當")
                DPerson.Visible = True
            Case 2  '修改
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '隱藏
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")
        'Slider Group Code
        Select Case FindFieldInf("SliderGRCode")
            Case 0  '顯示
                DSliderGRCode.BackColor = Color.PeachPuff
                DSliderGRCode.ReadOnly = True
                DSliderGRCode.Visible = True
            Case 1  '修改+檢查
                DSliderGRCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderGRCodeRqd", "DSliderGRCode", "異常：需輸入 Slider Code")
                DSliderGRCode.Visible = True
            Case 2  '修改
                DSliderGRCode.BackColor = Color.Yellow
                DSliderGRCode.Visible = True
            Case Else   '隱藏
                DSliderGRCode.Visible = False
        End Select
        If pPost = "New" Then DSliderGRCode.Text = ""
        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '顯示
                DSliderCode.Visible = False
                BSliderCode.Visible = False
                LSliderCode.Visible = True
            Case 1  '修改+檢查
                DSliderCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "異常：需輸入 Slider Code")
                DSliderCode.Visible = True
                BSliderCode.Visible = True
                LSliderCode.Visible = True
            Case 2  '修改
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.Visible = True
                BSliderCode.Visible = True
                LSliderCode.Visible = True
            Case Else   '隱藏
                DSliderCode.Visible = False
                BSliderCode.Visible = False
                LSliderCode.Visible = False
        End Select
        If pPost = "New" Then DSliderCode.Text = ""
        'Spec
        Select Case FindFieldInf("Spec")
            Case 0  '顯示
                DSpec.BackColor = Color.PeachPuff
                DSpec.ReadOnly = True
                DSpec.Visible = True
                BSpec.Visible = False
            Case 1  '修改+檢查
                DSpec.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSpecRqd", "DSpec", "異常：需輸入規格(Size, Chain Type, 胴體)")
                DSpec.Visible = True
                BSpec.Visible = True
            Case 2  '修改
                DSpec.BackColor = Color.Yellow
                DSpec.Visible = True
                BSpec.Visible = True
            Case Else   '隱藏
                DSpec.Visible = False
                BSpec.Visible = False
        End Select
        If pPost = "New" Then DSpec.Text = ""
        '圖號
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.PeachPuff
                DMapNo.ReadOnly = True
                DMapNo.Visible = True
                BOMapNo.Visible = False
                BMMapNo.Visible = False
            Case 1  '修改+檢查
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "異常：需輸入圖號")
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case 2  '修改
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case Else   '隱藏
                DMapNo.Visible = False
                BOMapNo.Visible = False
                BMMapNo.Visible = False
        End Select
        If pPost = "New" Then DMapNo.Text = ""

        '圖檔1
        Select Case FindFieldInf("MapFile1")
            Case 0  '顯示
                DMapFile1.Visible = False
                LMapFile1.Visible = True
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMapFile1Rqd", "DMapFile1", "異常：需輸入圖檔")
                DMapFile1.Visible = True
                LMapFile1.Visible = True
            Case 2  '修改
                DMapFile1.Visible = True
                LMapFile1.Visible = True
            Case Else   '隱藏
                DMapFile1.Visible = False
                LMapFile1.Visible = False
        End Select
        '圖檔2
        Select Case FindFieldInf("MapFile2")
            Case 0  '顯示
                DMapFile2.Visible = False
                LMapFile2.Visible = True
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMapFile2Rqd", "DMapFile2", "異常：需輸入圖檔")
                DMapFile2.Visible = True
                LMapFile2.Visible = True
            Case 2  '修改
                DMapFile2.Visible = True
                LMapFile2.Visible = True
            Case Else   '隱藏
                DMapFile2.Visible = False
                LMapFile2.Visible = False
        End Select
        '難易度
        Select Case FindFieldInf("Level")
            Case 0  '顯示
                DLevel.BackColor = Color.PeachPuff
                DLevel.Enabled = False
                DLevel.Visible = True
            Case 1  '修改+檢查
                DLevel.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLevelRqd", "DLevel", "異常：需輸入難易度")
                DLevel.Visible = True
            Case 2  '修改
                DLevel.BackColor = Color.Yellow
                DLevel.Visible = True
            Case Else   '隱藏
                DLevel.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")
        '組立判別
        Select Case FindFieldInf("Assembler")
            Case 0  '顯示
                DAssembler.BackColor = Color.PeachPuff
                DAssembler.ReadOnly = True
                DAssembler.Visible = True
            Case 1  '修改+檢查
                DAssembler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAssemblerRqd", "DAssembler", "異常：需輸入組立判定")
                DAssembler.Visible = True
            Case 2  '修改
                DAssembler.BackColor = Color.Yellow
                DAssembler.Visible = True
            Case Else   '隱藏
                DAssembler.Visible = False
        End Select
        If pPost = "New" Then DAssembler.Text = ""
        '拉頭種類(內製.外注...)
        Select Case FindFieldInf("SliderType1")
            Case 0  '顯示
                DSliderType1.BackColor = Color.PeachPuff
                DSliderType1.Enabled = False
                DSliderType1.Visible = True
            Case 1  '修改+檢查
                DSliderType1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderType1Rqd", "DSliderType1", "異常：需輸入拉頭種類")
                DSliderType1.Visible = True
            Case 2  '修改
                DSliderType1.BackColor = Color.Yellow
                DSliderType1.Visible = True
            Case Else   '隱藏
                DSliderType1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType1", "ZZZZZZ")
        '拉頭種類(半成品.成品...)
        Select Case FindFieldInf("SliderType2")
            Case 0  '顯示
                DSliderType2.BackColor = Color.PeachPuff
                DSliderType2.Enabled = False
                DSliderType2.Visible = True
            Case 1  '修改+檢查
                DSliderType2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderType2Rqd", "DSliderType2", "異常：需輸入拉頭種類")
                DSliderType2.Visible = True
            Case 2  '修改
                DSliderType2.BackColor = Color.Yellow
                DSliderType2.Visible = True
            Case Else   '隱藏
                DSliderType2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SliderType2", "ZZZZZZ")
        '生產地
        Select Case FindFieldInf("ManufPlace")
            Case 0  '顯示
                DManufPlace.BackColor = Color.PeachPuff
                DManufPlace.Enabled = False
                DManufPlace.Visible = True
            Case 1  '修改+檢查
                DManufPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufPlaceRqd", "DManufPlace", "異常：需輸入生產地")
                DManufPlace.Visible = True
            Case 2  '修改
                DManufPlace.BackColor = Color.Yellow
                DManufPlace.Visible = True
            Case Else   '隱藏
                DManufPlace.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ManufPlace", "ZZZZZZ")
        '材質
        Select Case FindFieldInf("Material")
            Case 0  '顯示
                DMaterial.BackColor = Color.PeachPuff
                DMaterial.Enabled = False
                DMaterial.Visible = True
            Case 1  '修改+檢查
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "異常：需輸入材質")
                DMaterial.Visible = True
            Case 2  '修改
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '隱藏
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")
        '材質細項
        Select Case FindFieldInf("MaterialDetail")
            Case 0  '顯示
                DMaterialDetail.BackColor = Color.PeachPuff
                DMaterialDetail.Enabled = False
                DMaterialDetail.Visible = True
            Case 1  '修改+檢查
                DMaterialDetail.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialDetailRqd", "DMaterialDetail", "異常：需輸入材質細項")
                DMaterialDetail.Visible = True
            Case 2  '修改
                DMaterialDetail.BackColor = Color.Yellow
                DMaterialDetail.Visible = True
            Case Else   '隱藏
                DMaterialDetail.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MaterialDetail", "ZZZZZZ")
        '材質備註
        Select Case FindFieldInf("MaterialOther")
            Case 0  '顯示
                DMaterialOther.BackColor = Color.PeachPuff
                DMaterialOther.ReadOnly = True
                DMaterialOther.Visible = True
            Case 1  '修改+檢查
                DMaterialOther.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialOtherRqd", "DMaterialOther", "異常：需輸入材質備註")
                DMaterialOther.Visible = True
            Case 2  '修改
                DMaterialOther.BackColor = Color.Yellow
                DMaterialOther.Visible = True
            Case Else   '隱藏
                DMaterialOther.Visible = False
        End Select
        If pPost = "New" Then DMaterialOther.Text = ""
        '委託廠商
        Select Case FindFieldInf("SellVendor")
            Case 0  '顯示
                DSellVendor.BackColor = Color.PeachPuff
                DSellVendor.ReadOnly = True
                DSellVendor.Visible = True
            Case 1  '修改+檢查
                DSellVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSellVendorRqd", "DSellVendor", "異常：需輸入委託廠商")
                DSellVendor.Visible = True
            Case 2  '修改
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.Visible = True
            Case Else   '隱藏
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""
        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.PeachPuff
                DBuyer.ReadOnly = True
                DBuyer.Visible = True
            Case 1  '修改+檢查
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
            Case 2  '修改
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '隱藏
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then DBuyer.Text = ""
        '確認書
        Select Case FindFieldInf("ConfirmFile")
            Case 0  '顯示
                DConfirmFile.Visible = False
                LConfirmFile.Visible = True
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DConfirmFileRqd", "DConfirmFile", "異常：需輸入確認書")
                DConfirmFile.Visible = True
                LConfirmFile.Visible = True
            Case 2  '修改
                DConfirmFile.Visible = True
                LConfirmFile.Visible = True
            Case Else   '隱藏
                DConfirmFile.Visible = False
                LConfirmFile.Visible = False
        End Select
        '授權書
        Select Case FindFieldInf("AuthorizeFile")
            Case 0  '顯示
                DAuthorizeFile.Visible = False
                LAuthorizeFile.Visible = True
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAuthorizeFileRqd", "DAuthorizeFile", "異常：需輸入授權書")
                DAuthorizeFile.Visible = True
                LAuthorizeFile.Visible = True
            Case 2  '修改
                DAuthorizeFile.Visible = True
                LAuthorizeFile.Visible = True
            Case Else   '隱藏
                DAuthorizeFile.Visible = False
                LAuthorizeFile.Visible = False
        End Select
        '開發理由
        Select Case FindFieldInf("DevReason")
            Case 0  '顯示
                DDevReason.BackColor = Color.PeachPuff
                DDevReason.ReadOnly = True
                DDevReason.Visible = True
            Case 1  '修改+檢查
                DDevReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevReasonRqd", "DDevReason", "異常：需輸入開發理由")
                DDevReason.Visible = True
            Case 2  '修改
                DDevReason.BackColor = Color.Yellow
                DDevReason.Visible = True
            Case Else   '隱藏
                DDevReason.Visible = False
        End Select
        If pPost = "New" Then DDevReason.Text = ""
        'Sample
        Select Case FindFieldInf("Sample")
            Case 0  '顯示
                DSample.Visible = False
                BSample.Visible = False
                LSample.Visible = True
            Case 1  '修改+檢查
                DSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleRqd", "DSample", "異常：需輸入樣品資訊")
                DSample.Visible = True
                BSample.Visible = True
                LSample.Visible = True
            Case 2  '修改
                DSample.BackColor = Color.Yellow
                DSample.Visible = True
                BSample.Visible = True
                LSample.Visible = True
            Case Else   '隱藏
                DSample.Visible = False
                BSample.Visible = False
                LSample.Visible = False
        End Select
        If pPost = "New" Then DSample.Text = ""
        '樣品圖
        Select Case FindFieldInf("SampleFile")
            Case 0  '顯示
                DSampleFile.Visible = False
                LSampleFile.Visible = True
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "異常：需輸入樣品圖")
                DSampleFile.Visible = True
                LSampleFile.Visible = True
            Case 2  '修改
                DSampleFile.Visible = True
                LSampleFile.Visible = True
            Case Else   '隱藏
                DSampleFile.Visible = False
                LSampleFile.Visible = False
        End Select
        'Price
        Select Case FindFieldInf("Price")
            Case 0  '顯示
                DPrice.Visible = False
                BPrice.Visible = False
                LPrice.Visible = True
            Case 1  '修改+檢查
                DPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPriceRqd", "DPrice", "異常：需輸入單價資訊")
                DPrice.Visible = True
                BPrice.Visible = True
                LPrice.Visible = True
            Case 2  '修改
                DPrice.BackColor = Color.Yellow
                DPrice.Visible = True
                BPrice.Visible = True
                LPrice.Visible = True
            Case Else   '隱藏
                DPrice.Visible = False
                BPrice.Visible = False
                LPrice.Visible = False
        End Select
        If pPost = "New" Then DPrice.Text = ""
        '應收模具費
        Select Case FindFieldInf("ArMoldFee")
            Case 0  '顯示
                DArMoldFee.BackColor = Color.PeachPuff
                DArMoldFee.ReadOnly = True
                DArMoldFee.Visible = True
            Case 1  '修改+檢查
                DArMoldFee.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DArMoldFeeRqd", "DArMoldFee", "異常：需輸入應收模具費")
                DArMoldFee.Visible = True
            Case 2  '修改
                DArMoldFee.BackColor = Color.Yellow
                DArMoldFee.Visible = True
            Case Else   '隱藏
                DArMoldFee.Visible = False
        End Select
        If pPost = "New" Then DArMoldFee.Text = ""
        '模具購入價
        Select Case FindFieldInf("PurMold")
            Case 0  '顯示
                DPurMold.BackColor = Color.PeachPuff
                DPurMold.ReadOnly = True
                DPurMold.Visible = True
            Case 1  '修改+檢查
                DPurMold.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPurMoldRqd", "DPurMold", "異常：需輸入模具購入價")
                DPurMold.Visible = True
            Case 2  '修改
                DPurMold.BackColor = Color.Yellow
                DPurMold.Visible = True
            Case Else   '隱藏
                DPurMold.Visible = False
        End Select
        If pPost = "New" Then DPurMold.Text = ""
        '引手購入單價
        Select Case FindFieldInf("PullerPrice")
            Case 0  '顯示
                DPullerPrice.BackColor = Color.PeachPuff
                DPullerPrice.ReadOnly = True
                DPullerPrice.Visible = True
            Case 1  '修改+檢查
                DPullerPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPullerPriceRqd", "DPullerPrice", "異常：需輸入引手購入單價")
                DPullerPrice.Visible = True
            Case 2  '修改
                DPullerPrice.BackColor = Color.Yellow
                DPullerPrice.Visible = True
            Case Else   '隱藏
                DPullerPrice.Visible = False
        End Select
        If pPost = "New" Then DPullerPrice.Text = ""
        '模具名稱
        Select Case FindFieldInf("MoldName")
            Case 0  '顯示
                DMoldName.BackColor = Color.PeachPuff
                DMoldName.ReadOnly = True
                DMoldName.Visible = True
            Case 1  '修改+檢查
                DMoldName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldNameRqd", "DMoldName", "異常：需輸入模具名稱")
                DMoldName.Visible = True
            Case 2  '修改
                DMoldName.BackColor = Color.Yellow
                DMoldName.Visible = True
            Case Else   '隱藏
                DMoldName.Visible = False
        End Select
        If pPost = "New" Then DMoldName.Text = ""
        '模型個取-模型
        Select Case FindFieldInf("MoldQty")
            Case 0  '顯示
                DMoldQty.BackColor = Color.PeachPuff
                DMoldQty.ReadOnly = True
                DMoldQty.Visible = True
            Case 1  '修改+檢查
                DMoldQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldQtyRqd", "DMoldQty", "異常：需輸入模型個取-模型")
                DMoldQty.Visible = True
            Case 2  '修改
                DMoldQty.BackColor = Color.Yellow
                DMoldQty.Visible = True
            Case Else   '隱藏
                DMoldQty.Visible = False
        End Select
        If pPost = "New" Then DMoldQty.Text = ""
        '模型個取-穴取
        Select Case FindFieldInf("MoldPoint")
            Case 0  '顯示
                DMoldPoint.BackColor = Color.PeachPuff
                DMoldPoint.ReadOnly = True
                DMoldPoint.Visible = True
            Case 1  '修改+檢查
                DMoldPoint.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldPointRqd", "DMoldPoint", "異常：需輸入模型個取-穴取")
                DMoldPoint.Visible = True
            Case 2  '修改
                DMoldPoint.BackColor = Color.Yellow
                DMoldPoint.Visible = True
            Case Else   '隱藏
                DMoldPoint.Visible = False
        End Select
        If pPost = "New" Then DMoldPoint.Text = ""
        'Quality1
        Select Case FindFieldInf("Quality1")
            Case 0  '顯示
                DQuality1.Visible = False
                BQuality1.Visible = False
                LQuality1.Visible = True
            Case 1  '修改+檢查
                DQuality1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQuality1Rqd", "DQuality1", "異常：需輸入品質分析資訊")
                DQuality1.Visible = True
                BQuality1.Visible = True
                LQuality1.Visible = True
            Case 2  '修改
                DQuality1.BackColor = Color.Yellow
                DQuality1.Visible = True
                BQuality1.Visible = True
                LQuality1.Visible = True
            Case Else   '隱藏
                DQuality1.Visible = False
                BQuality1.Visible = False
                LQuality1.Visible = False
        End Select
        If pPost = "New" Then DQuality1.Text = ""
        'Quality2
        Select Case FindFieldInf("Quality2")
            Case 0  '顯示
                DQuality2.Visible = False
                BQuality2.Visible = False
                LQuality2.Visible = True
            Case 1  '修改+檢查
                DQuality2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQuality2Rqd", "DQuality2", "異常：需輸入品質分析資訊")
                DQuality2.Visible = True
                BQuality2.Visible = True
                LQuality2.Visible = True
            Case 2  '修改
                DQuality2.BackColor = Color.Yellow
                DQuality2.Visible = True
                BQuality2.Visible = True
                LQuality2.Visible = True
            Case Else   '隱藏
                DQuality2.Visible = False
                BQuality2.Visible = False
                LQuality2.Visible = False
        End Select
        If pPost = "New" Then DQuality2.Text = ""
        '品質分析書
        Select Case FindFieldInf("QAAttachFile")
            Case 0  '顯示
                DQAAttachFile.Visible = False
                LQAAttachFile.Visible = True
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DQAAttachFileRqd", "DQAAttachFile", "異常：需輸入品質分析書")
                DQAAttachFile.Visible = True
                LQAAttachFile.Visible = True
            Case 2  '修改
                DQAAttachFile.Visible = True
                LQAAttachFile.Visible = True
            Case Else   '隱藏
                DQAAttachFile.Visible = False
                LQAAttachFile.Visible = False
        End Select
        '單號
        If pPost = "New" Then DFormSno.Text = ""
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top + 25))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '部門
        If pFieldName = "Division" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='DIVISION' Order by Data "
            DDivision.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DDivision.Items.Add(ListItem1)
            Next
        End If
        '擔當
        If pFieldName = "Person" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='PERSON' Order by Data "
            DPerson.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DPerson.Items.Add(ListItem1)
            Next
        End If
        '難易度
        If pFieldName = "Level" Then
            SQL = "Select * From M_Referp Where Cat='007' and DKey<>'Z' Order by DKey, Data "
            DLevel.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DLevel.Items.Add(ListItem1)
            Next
        End If
        '拉頭種類(內製.外注...)
        If pFieldName = "SliderType1" Then
            SQL = "Select * From M_Referp Where Cat='200' and DKey='SLIDERTYPE1' Order by Data "
            DSliderType1.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DSliderType1.Items.Add(ListItem1)
            Next
        End If
        '拉頭種類(半成品.成品...)
        If pFieldName = "SliderType2" Then
            SQL = "Select * From M_Referp Where Cat='200' and DKey='SLIDERTYPE2' Order by Data "
            DSliderType2.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DSliderType2.Items.Add(ListItem1)
            Next
        End If
        '生產地
        If pFieldName = "ManufPlace" Then
            SQL = "Select * From M_Referp Where Cat='200' and DKey='MANUFPLACE' Order by Data "
            DManufPlace.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DManufPlace.Items.Add(ListItem1)
            Next
        End If
        '材質
        If pFieldName = "Material" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='MATERIAL' Order by Data "
            DMaterial.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DMaterial.Items.Add(ListItem1)
            Next
        End If
        '材質細項
        If pFieldName = "MaterialDetail" Then
            SQL = "Select * From M_Referp Where Cat='101' and DKey= '" & DMaterial.SelectedValue & "' Order by Data "
            DMaterialDetail.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DMaterialDetail.Items.Add(ListItem1)
            Next
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     尋找表單欄位屬性
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i < 44 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     材質點選後事件
    '**
    '*****************************************************************
    Private Sub DMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMaterial.SelectedIndexChanged
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        SQL = "Select * From M_Referp Where Cat='101' and DKey= '" & DMaterial.SelectedValue & "' Order by Data "
        DMaterialDetail.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1        '顯示細項對應內容
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DMaterialDetail.Items.Add(ListItem1)
        Next
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        '上傳日期
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)

        Dim NewFormSno As Integer = wFormSno    '表單流水號
        Dim RtnCode As Integer = 0
        '--其他SubFile---------
        Dim oPutSubFile As Object
        '--取得表單流水號設定---------
        Dim oGetSeqNo As Object

        'Check附圖-1Size及格式
        If ErrCode = 0 Then
            If DMapFile1.Visible Then
                If DMapFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DMapFile1)
                End If
            End If
        End If
        'Check附圖-2Size及格式
        If ErrCode = 0 Then
            If DMapFile2.Visible Then
                If DMapFile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DMapFile2)
                End If
            End If
        End If
        'Check確認書Size及格式
        If ErrCode = 0 Then
            If DConfirmFile.Visible Then
                If DConfirmFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DConfirmFile)
                End If
            End If
        End If
        'Check授權書Size及格式
        If ErrCode = 0 Then
            If DAuthorizeFile.Visible Then
                If DAuthorizeFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DAuthorizeFile)
                End If
            End If
        End If
        'Check樣品圖Size及格式
        If ErrCode = 0 Then
            If DSampleFile.Visible Then
                If DSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DSampleFile)
                End If
            End If
        End If
        'CheckQA附件Size及格式
        If ErrCode = 0 Then
            If DQAAttachFile.Visible Then
                If DQAAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DQAAttachFile)
                End If
            End If
        End If

        '儲存資料
        If ErrCode = 0 Then
            Dim Path As String
            If wOFormNo = "000003" Then
                Path = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInFilePath"))
            Else
                Path = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufOutFilePath"))
            End If
            Dim FileName As String
            Dim SQL As String

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            Dim OleDBCommand1 As New OleDbCommand

            '儲存表單資料
            If wFormSno > 0 Then    '判斷是否有資料
                SQL = "Update F_ModManufSheet Set "
                SQL = SQL + " No = '" & DNo.Text & "',"
                SQL = SQL + " Date = '" & DDate.Text & "',"
                SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
                SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
                If DSliderCode.Text <> "" Then
                    SQL = SQL + " SliderCode = '" & "1" & "',"
                Else
                    SQL = SQL + " SliderCode = '" & "0" & "',"
                End If
                SQL = SQL + " SliderGRCode = '" & DSliderGRCode.Text & "',"
                SQL = SQL + " Spec = '" & DSpec.Text & "',"
                SQL = SQL + " MapNo = '" & DMapNo.Text & "',"

                If DMapFile1.Visible Then
                    If DMapFile1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        FileName = CStr(wFormSno) & "-" & "Map1" & "-" & UploadDateTime & "-" & Right(DMapFile1.PostedFile.FileName, InStr(StrReverse(DMapFile1.PostedFile.FileName), "\") - 1)
                        Try    '上傳圖檔
                            DMapFile1.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " MapFile1 = '" & FileName & "',"
                    End If
                End If

                If DMapFile2.Visible Then
                    If DMapFile2.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        FileName = CStr(wFormSno) & "-" & "Map2" & "-" & UploadDateTime & "-" & Right(DMapFile2.PostedFile.FileName, InStr(StrReverse(DMapFile2.PostedFile.FileName), "\") - 1)
                        Try    '上傳圖檔
                            DMapFile2.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " MapFile2 = '" & FileName & "',"
                    End If
                End If

                SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
                SQL = SQL + " Assembler = '" & DAssembler.Text & "',"
                SQL = SQL + " SliderType1 = '" & DSliderType1.SelectedValue & "',"
                SQL = SQL + " SliderType2 = '" & DSliderType2.SelectedValue & "',"
                SQL = SQL + " ManufPlace = '" & DManufPlace.SelectedValue & "',"
                SQL = SQL + " Material = '" & DMaterial.SelectedValue & "',"
                SQL = SQL + " MaterialDetail = '" & DMaterialDetail.SelectedValue & "',"
                SQL = SQL + " MaterialOther = '" & DMaterialOther.Text & "',"
                SQL = SQL + " SellVendor = '" & DSellVendor.Text & "',"
                SQL = SQL + " Buyer = '" & DBuyer.Text & "',"

                If DConfirmFile.Visible Then
                    If DConfirmFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        FileName = CStr(wFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                        Try    '上傳圖檔
                            DConfirmFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " ConfirmFile = '" & FileName & "',"
                    End If
                End If

                If DAuthorizeFile.Visible Then
                    If DAuthorizeFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        FileName = CStr(wFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                        Try    '上傳圖檔
                            DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " AuthorizeFile = '" & FileName & "',"
                    End If
                End If

                SQL = SQL + " DevReason = '" & DDevReason.Text & "',"
                If DSample.Text <> "" Then
                    SQL = SQL + " Sample = '" & "1" & "',"
                Else
                    SQL = SQL + " Sample = '" & "0" & "',"
                End If
                If DPrice.Text <> "" Then
                    SQL = SQL + " Price = '" & "1" & "',"
                Else
                    SQL = SQL + " Price = '" & "0" & "',"
                End If
                SQL = SQL + " ArMoldFee = '" & DArMoldFee.Text & "',"
                SQL = SQL + " PurMold = '" & DPurMold.Text & "',"
                SQL = SQL + " PullerPrice = '" & DPullerPrice.Text & "',"
                SQL = SQL + " MoldName = '" & DMoldName.Text & "',"
                SQL = SQL + " MoldQty = '" & DMoldQty.Text & "',"
                SQL = SQL + " MoldPoint = '" & DMoldPoint.Text & "',"
                If DQuality1.Text <> "" Then
                    SQL = SQL + " Quality1 = '" & "1" & "',"
                Else
                    SQL = SQL + " Quality1 = '" & "0" & "',"
                End If
                If DQuality2.Text <> "" Then
                    SQL = SQL + " Quality2 = '" & "1" & "',"
                Else
                    SQL = SQL + " Quality2 = '" & "0" & "',"
                End If

                If DQAAttachFile.Visible Then
                    If DQAAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAAttachFile.PostedFile.FileName, InStr(StrReverse(DQAAttachFile.PostedFile.FileName), "\") - 1)
                        Try    '上傳圖檔
                            DQAAttachFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " QAAttachFile = '" & FileName & "',"
                    End If
                End If

                If DSampleFile.Visible Then
                    If DSampleFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        FileName = CStr(wFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                        Try    '上傳圖檔
                            DSampleFile.PostedFile.SaveAs(Path + FileName)
                        Catch ex As Exception
                        End Try
                        SQL = SQL + " SampleFile = '" & FileName & "',"
                    End If
                End If
                SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQL
                OleDbConnection1.Open()
                OleDBCommand1.ExecuteNonQuery()
                OleDbConnection1.Close()

                'Slider Table處理
                oPutSubFile = Server.CreateObject("MSubFile.SubFile")
                RtnCode = oPutSubFile.PutSlider(DSliderCode.Text, DSliderGRCode.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'Sample Table處理
                RtnCode = oPutSubFile.PutSample(DSample.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'Price Table處理
                RtnCode = oPutSubFile.PutPrice(DPrice.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'QA1 Table處理
                RtnCode = oPutSubFile.PutQA1(DQuality1.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
                'QA2 Table處理
                RtnCode = oPutSubFile.PutQA2(DQuality2.Text, "900003", wFormSno, Request.Cookies("UserID").Value)
            Else
                '取得表單流水號
                oGetSeqNo = Server.CreateObject("GetSeqno.WFFormInf")
                RtnCode = oGetSeqNo.Seqno(wFormNo, NewFormSno)   '表單號碼, 表單流水號
                If RtnCode <> 0 Then
                    ErrCode = 9110
                Else
                    SQL = "Insert into F_ModManufSheet "
                    SQL = SQL + "( "
                    SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, "                         '1~4
                    SQL = SQL + "No, Date, Division, Person, SliderCode, "                      '5~9
                    SQL = SQL + "SliderGRCode, Spec, MapNo, MapFile1, MapFile2, "               '10~14
                    SQL = SQL + "Level, Assembler, SliderType1, SliderType2, ManufPlace, "      '15~19
                    SQL = SQL + "Material, MaterialDetail, MaterialOther, SellVendor, Buyer, "  '20~24
                    SQL = SQL + "ConfirmFile, AuthorizeFile, DevReason, Sample, Price, "        '25~29
                    SQL = SQL + "ArMoldFee, PurMold, PullerPrice, MoldName, MoldQty, "          '29~33
                    SQL = SQL + "MoldPoint, Quality1, Quality2, QAAttachFile, SampleFile,  "    '34~38
                    SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "               '39~42
                    SQL = SQL + ")  "
                    SQL = SQL + "VALUES( "
                    SQL = SQL + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
                    SQL = SQL + " '" + NowDateTime + "', "              '結案日
                    SQL = SQL + " '900003', "                           '表單代號
                    SQL = SQL + " '" + Str(NewFormSno) + "', "            '表單流水號

                    SQL = SQL + " '" + DNo.Text + "', "                 'NO
                    SQL = SQL + " '" + DDate.Text + "', "               '日期
                    SQL = SQL + " '" + DDivision.SelectedValue + "', "  '部門
                    SQL = SQL + " '" + DPerson.SelectedValue + "', "    '擔當
                    If DSliderCode.Text <> "" Then                      'Slider Code(顯示用)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    SQL = SQL + " '" + DSliderGRCode.Text + "', "       'Slider Group Code
                    SQL = SQL + " '" + DSpec.Text + "', "               '規格
                    SQL = SQL + " '" + DMapNo.Text + "', "              '圖號

                    If DMapFile1.Visible Then                           '圖檔1
                        If DMapFile1.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                            FileName = CStr(wFormSno) & "-" & "Map1" & "-" & UploadDateTime & "-" & Right(DMapFile1.PostedFile.FileName, InStr(StrReverse(DMapFile1.PostedFile.FileName), "\") - 1)
                            Try    '上傳圖檔
                                DMapFile1.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wMapFile1.Text
                        End If
                    Else
                        FileName = wMapFile1.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    If DMapFile2.Visible Then                           '圖檔2
                        If DMapFile2.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                            FileName = CStr(wFormSno) & "-" & "Map2" & "-" & UploadDateTime & "-" & Right(DMapFile2.PostedFile.FileName, InStr(StrReverse(DMapFile2.PostedFile.FileName), "\") - 1)
                            Try    '上傳圖檔
                                DMapFile2.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wMapFile2.Text
                        End If
                    Else
                        FileName = wMapFile2.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    SQL = SQL + " '" + DLevel.SelectedValue + "', "           '難易度
                    SQL = SQL + " '" + DAssembler.Text + "', "                '組立判定
                    SQL = SQL + " '" + DSliderType1.SelectedValue + "', "     '拉頭種類-1
                    SQL = SQL + " '" + DSliderType2.SelectedValue + "', "     '拉頭種類-2
                    SQL = SQL + " '" + DManufPlace.SelectedValue + "', "      '生產地
                    SQL = SQL + " '" + DMaterial.SelectedValue + "', "        '材質
                    SQL = SQL + " '" + DMaterialDetail.SelectedValue + "', "  '材質細項
                    SQL = SQL + " '" + DMaterialOther.Text + "', "            '材質其他說明
                    SQL = SQL + " '" + DSellVendor.Text + "', "               '委託廠商
                    SQL = SQL + " '" + DBuyer.Text + "', "                    'Buyer
                    If DConfirmFile.Visible Then                              '確認書
                        If DConfirmFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                            FileName = CStr(wFormSno) & "-" & "Confirm" & "-" & UploadDateTime & "-" & Right(DConfirmFile.PostedFile.FileName, InStr(StrReverse(DConfirmFile.PostedFile.FileName), "\") - 1)
                            Try    '上傳圖檔
                                DConfirmFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wConfirmFile.Text
                        End If
                    Else
                        FileName = wConfirmFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    If DAuthorizeFile.Visible Then                            '授權書
                        If DAuthorizeFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                            FileName = CStr(wFormSno) & "-" & "Authorize" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                            Try    '上傳圖檔
                                DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wAuthorizeFile.Text
                        End If
                    Else
                        FileName = wAuthorizeFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    SQL = SQL + " '" + DDevReason.Text + "', "          '開發理由
                    If DSample.Text <> "" Then                          'Sample(顯示用)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    If DPrice.Text <> "" Then                           'Price(顯示用)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    SQL = SQL + " '" + DArMoldFee.Text + "', "          '應收模具費
                    SQL = SQL + " '" + DPurMold.Text + "', "            '模具購入費
                    SQL = SQL + " '" + DPullerPrice.Text + "', "        '引手購入單價
                    SQL = SQL + " '" + DMoldName.Text + "', "           '模具名稱
                    SQL = SQL + " '" + DMoldQty.Text + "', "            '模型數
                    SQL = SQL + " '" + DMoldPoint.Text + "', "          '穴取數
                    If DQuality1.Text <> "" Then                        '品質分析-1(顯示用)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If
                    If DQuality2.Text <> "" Then                        '品質分析-2(顯示用)
                        SQL = SQL + " '1', "
                    Else
                        SQL = SQL + " '0', "
                    End If

                    If DQAAttachFile.Visible Then                      '品質分析附件
                        If DQAAttachFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                            FileName = CStr(wFormSno) & "-" & "QA" & "-" & UploadDateTime & "-" & Right(DQAAttachFile.PostedFile.FileName, InStr(StrReverse(DQAAttachFile.PostedFile.FileName), "\") - 1)
                            Try    '上傳圖檔
                                DQAAttachFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wQAAttachFile.Text
                        End If
                    Else
                        FileName = wQAAttachFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    If DSampleFile.Visible Then                      '樣品圖檔
                        If DSampleFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                            FileName = CStr(wFormSno) & "-" & "Sample" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                            Try    '上傳圖檔
                                DSampleFile.PostedFile.SaveAs(Path + FileName)
                            Catch ex As Exception
                            End Try
                        Else
                            FileName = wSampleFile.Text
                        End If
                    Else
                        FileName = wSampleFile.Text
                    End If
                    SQL = SQL + " '" + FileName + "', "

                    SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '作成者
                    SQL = SQL + " '" + NowDateTime + "', "       '作成時間
                    SQL = SQL + " '" + "" + "', "                       '修改者
                    SQL = SQL + " '" + NowDateTime + "' "       '修改時間
                    SQL = SQL + " ) "
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQL
                    OleDbConnection1.Open()
                    OleDBCommand1.ExecuteNonQuery()
                    OleDbConnection1.Close()

                    'Slider Table處理
                    oPutSubFile = Server.CreateObject("MSubFile.SubFile")
                    RtnCode = oPutSubFile.PutSlider(DSliderCode.Text, DSliderGRCode.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'Sample Table處理
                    RtnCode = oPutSubFile.PutSample(DSample.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'Price Table處理
                    RtnCode = oPutSubFile.PutPrice(DPrice.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'QA1 Table處理
                    RtnCode = oPutSubFile.PutQA1(DQuality1.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                    'QA2 Table處理
                    RtnCode = oPutSubFile.PutQA2(DQuality2.Text, "900003", NewFormSno, Request.Cookies("UserID").Value)
                End If
            End If
        Else
            If ErrCode = 9110 Then Message = "取得表單流水號計算異常,請確認或連絡系統人員!"
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0

        If ErrCode = 0 Then
            If wFormSno > 0 And wStep > 1 Then    '判斷是否[簽核]
                Response.Write(String.Format("<script>window.close();</script>"))
            Else
                Dim SQL As String
                Dim DBDataSet1 As New DataSet
                'DB連結設定
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                OleDbConnection1.Open()

                SQL = "Select FormNo, FormSno From F_ModManufSheet "
                SQL = SQL & " Where Sts = 0 "
                SQL = SQL & "   And FormNo =  '900003' "
                SQL = SQL & "   And FormSno =  '" & CStr(NewFormSno) & "'"
                SQL = SQL & " Order by Unique_ID Desc "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
                If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
                    rFormNo = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("FormNo")
                    rFormSno = DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("FormSno")
                End If
                'DB連結關閉
                OleDbConnection1.Close()
                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();</script>", "Form1.DNFormNo", rFormNo, "Form1.DNFormSno", rFormSno))
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".jpg", ".jpeg", ".gif", ".xls", ".doc"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If

        'If UPFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
        'Check上傳檔案格式
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function

End Class
