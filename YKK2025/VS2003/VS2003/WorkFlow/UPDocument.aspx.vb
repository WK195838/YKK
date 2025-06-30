Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class UPDocument
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DUPDocument As System.Web.UI.WebControls.Image
    Protected WithEvents DType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents DYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BMakerDate As System.Web.UI.WebControls.Button
    Protected WithEvents DClass As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BSave As System.Web.UI.WebControls.Button
    Protected WithEvents DMaker As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDescriptionRqd As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents DMakerTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDocFileName As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DDocFileNameRqd As System.Web.UI.WebControls.RequiredFieldValidator

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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "UPDocument.aspx"

        If Not Me.IsPostBack Then   '不是PostBack
            CheckAuthority()
            SetInitData()          '設定畫面初始值
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定畫面初始值
    '**
    '*****************************************************************
    Sub SetInitData()
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '文件類別
        SQL = "Select * From M_Referp Where Cat='015' and DKey='CLASS' Order by Data "
        DClass.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DClass.Items.Add(ListItem1)
        Next

        '作成者
        DBDataSet1.Clear()
        SQL = "Select * From M_Referp Where Cat='015' and DKey='MAKER' Order by Data "
        DMaker.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DMaker.Items.Add(ListItem1)
        Next
        'DB連結關閉
        OleDbConnection1.Close()

        '年
        DYear.Items.Clear()
        For i = 2004 To 2020
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If ListItem1.Value = CStr(DateTime.Now.Year) Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next

        '月
        DMonth.Items.Clear()
        For i = 1 To 12
            Dim ListItem1 As New ListItem
            If i < 10 Then
                ListItem1.Text = "0" & CStr(i)
                ListItem1.Value = "0" & CStr(i)
            Else
                ListItem1.Text = CStr(i)
                ListItem1.Value = CStr(i)
            End If
            If CStr(i) = CStr(DateTime.Now.Month) Then ListItem1.Selected = True
            DMonth.Items.Add(ListItem1)
        Next

        DDescription.Text = ""  '說明
        DMakerTime.Text = CStr(DateTime.Now.Today)  '作成日
        BMakerDate.Attributes("onclick") = "calendarPicker('Form1.DMakerTime');"  '日期選擇
    End Sub

    Private Sub BSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""

        'Check上傳檔案Size及格式
        If ErrCode = 0 Then
            If DDocFileName.Visible Then
                If DDocFileName.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DDocFileName)
                End If
            End If
        End If

        If ErrCode = 0 Then
            SaveData()
            Response.Write(YKK.ShowMessage("檔案上傳及資料儲存成功"))
            SetInitData()          '設定畫面初始值
        Else
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '儲存表單ErrCode=0

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存資料
    '**
    '*****************************************************************
    Sub SaveData()
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("DocFilePath"))
        Dim FileName As String
        Dim SQL As String
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)    '上傳日期
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)                '現在日時


        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQL = "Insert into F_UPDocument "
        SQL = SQL + "( "
        SQL = SQL + "Sts, Class, Type, Year, Month, "                      '1~5
        SQL = SQL + "DocFileName, Description, Maker, MakerTime, "         '6~9
        SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "      '10~13
        SQL = SQL + ")  "
        SQL = SQL + "VALUES( "
        SQL = SQL + " '1', "                                 '狀態(0:未結,1:已結NG,2:已結OK)
        SQL = SQL + " '" + DClass.SelectedValue + "', "      '類別
        SQL = SQL + " '" + DType.SelectedValue + "', "       '性質
        SQL = SQL + " '" + DYear.SelectedValue + "', "       '年
        If DType.SelectedValue = "0" Then
            SQL = SQL + " '" + DMonth.SelectedValue + "', "      '月
        Else
            SQL = SQL + " '', "      '月
        End If
        If DDocFileName.Visible Then
            If DDocFileName.PostedFile.FileName <> "" Then   '上傳檔案
                FileName = "Doc" & "-" & UploadDateTime & "-" & Right(DDocFileName.PostedFile.FileName, InStr(StrReverse(DDocFileName.PostedFile.FileName), "\") - 1)
                Try
                    DDocFileName.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQL = SQL + " '" + FileName + "', "                 '上傳檔案名
        SQL = SQL + " '" + DDescription.Text + "', "        '簡介
        SQL = SQL + " '" + DMaker.SelectedValue + "', "     '作成者
        SQL = SQL + " '" + DMakerTime.Text + "', "          '作成日
        SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "  '作成者
        SQL = SQL + " '" + NowDateTime + "', "                      '作成時間
        SQL = SQL + " '" + "" + "', "                               '修改者
        SQL = SQL + " '" + NowDateTime + "' "                       '修改時間
        SQL = SQL + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".doc", ".xls"}   '定義允許的檔案格式
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
    End Function

End Class
