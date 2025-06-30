Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class AppendSpecSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DAppendSpecSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPriceA1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceA2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceA3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceB2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceB3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceC2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceC3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceD2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceD3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceE2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceE3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceF2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceG2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceG3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceH2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceH3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceI2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceI3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPriceJ2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPriceJ3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQAFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAssembler As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAssemblerFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DCPSC As System.Web.UI.WebControls.TextBox

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

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "AppendSpecSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("AppendSpecFilePath")
        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("AppendSpecFilePath")
        Dim RtnCode As Integer = 0

        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_AppendSpecSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_AppendSpecSheet")
        If DBDataSet1.Tables("F_AppendSpecSheet").Rows.Count > 0 Then

            DNo.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Person"))      '擔當
            DBuyer.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("SellVendor")       '委託廠商
            DAssembler.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Assembler")       '組立判定
            DCPSC.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("CPSC")       '組立判定
            DSliderCode.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("SliderCode")       'Slider Code
            DMapNo.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("MapNo")                 '圖號

            DOFormNo.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("OFormNo")             '原表單
            DOFormSno.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("OFormSno")           '原單號
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000011" Then
                    LOFormNo.NavigateUrl = "ImportSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    If DOFormNo.Text = "000003" Then
                        LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                    Else
                        LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                    End If
                End If

                DOFormNo.Visible = False
                DOFormSno.Visible = False
            Else
                LOFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Target")           '目的
            DContent.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Content")         '內容
            DDescription.Text = DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Description") 'EA說明

            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("Price") = 1 Then               '單價
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

            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("QAFile") <> "" Then          '測試報告


                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("QAFile"))

                LQAFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("QAFile")
            Else
                LQAFile.Visible = False
            End If


            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("AssemblerFile") <> "" Then          '組立報告書

                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("AssemblerFile"))

                LAssemblerFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("AssemblerFile")
            Else
                LAssemblerFile.Visible = False
            End If


            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("ContactFile") <> "" Then     '切結書 

                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("ContactFile"))

                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If
            If DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("RefFile") <> "" Then          '其他附件 
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("RefFile"))

                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_AppendSpecSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
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
    End Sub

End Class
