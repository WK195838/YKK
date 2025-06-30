Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SCD_SampleSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOP6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCDITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCSITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTNRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCNITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTNLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTHCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DECOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTACOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDEVPRD As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDEVNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTAWIDTH As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSAMPLEFILE As System.Web.UI.WebControls.Image
    Protected WithEvents DCODENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSIZENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTALINE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSampleSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSampleSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents DDATE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAPPBUYER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF5Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF6Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF7Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents WF2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF4Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF3Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF7 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF7Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF6Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF5Name As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO1Item As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO2Item As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOther1 As System.Web.UI.WebControls.Label
    Protected WithEvents DOther2 As System.Web.UI.WebControls.Label
    Protected WithEvents LQCFILE1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE5 As System.Web.UI.WebControls.HyperLink

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
        Response.Cookies("PGM").Value = "SCD_SampleSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SampleFilePath")
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_SampleSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SampleSheet")
        If DBDataSet1.Tables("F_SampleSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("No")                     'No
            DAPPBUYER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("AppBuyer")         'Customer
            DDATE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Date")                 '發行日
            DSIZENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SizeNo")             'Size
            DITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Item")                 'Item
            DCODENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CodeNo")             'Code No
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile") <> "" Then          '實際樣品
                LSAMPLEFILE.ImageUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile")
            Else
                LSAMPLEFILE.Visible = False
            End If
            DTAWIDTH.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TAWidth")           '布帶寬度
            DDEVNO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevNo")               '開發No
            DDEVPRD.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevPrd")             '開發期間
            DTACOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TACol")               '布帶Color
            DTALINE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TALine")             '條紋線Color
            DECOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("ECol")                 '務齒
            DCCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CCol")                 '丸紐
            DTHCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("THCol")               '縫工線
            DOTHER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other")               '其他

            '980410 update by alin
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1") <> "" Then              '品測報告1
                LQCFILE1.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1")
            Else
                LQCFILE1.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2") <> "" Then              '品測報告2
                LQCFILE2.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2")
            Else
                LQCFILE2.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3") <> "" Then              '品測報告3
                LQCFILE3.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3")
            Else
                LQCFILE3.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4") <> "" Then              '品測報告4
                LQCFILE4.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4")
            Else
                LQCFILE4.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5") <> "" Then              '品測報告5
                LQCFILE5.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5")
            Else
                LQCFILE5.Visible = False
            End If

            DOther1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other1")           'Other1
            DOther2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("other2")           'Other2
            DO1Item.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O1Item")           'O1Item
            DO2Item.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O2item")           'O2Item
            '=====================


            DTNLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNLItem")           'TAPE NAT-左
            DTNRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNRItem")           'TAPE NAT-右
            DTSLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSLItem")           'TAPE SET-左
            DTSRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSRItem")           'TAPE SET-右
            DTDLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDLItem")           'TAPE DYED-左
            DTDRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDRItem")           'TAPE DYED-右
            DCNITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CNItem")             'CHAIN NAT
            DCSITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CSItem")             'CHAIN SET
            DCDITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CDItem")             'CHAIN DYED
            DCITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CItem")               'CORD
            DOP1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP1")                   '工程1
            DOP2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP2")                   '工程2
            DOP3.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP3")                   '工程3
            DOP4.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP4")                   '工程4
            DOP5.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP5")                   '工程5
            DOP6.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP6")                   '工程6
            DWF1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF1")                   '承認WF1
            DWF2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF2")                   '承認WF2
            DWF3.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3")                   '承認WF3
            DWF4.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4")                   '承認WF4
            DWF5.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5")                   '承認WF5
            DWF6.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6")                   '承認WF6
            DWF7.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7")                   '承認WF7
            DWF3Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3Name")           '承認者部門WF3
            DWF4Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4Name")           '承認者部門WF4
            DWF5Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5Name")           '承認者部門WF5
            DWF6Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6Name")           '承認者部門WF6
            DWF7Name.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7Name")           '承認者部門WF7
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

End Class
