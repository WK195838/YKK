Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class UPDocument
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
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

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     �ۭq�[���w
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
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

        If Not Me.IsPostBack Then   '���OPostBack
            CheckAuthority()
            SetInitData()          '�]�w�e����l��
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�e����l��
    '**
    '*****************************************************************
    Sub SetInitData()
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        '������O
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

        '�@����
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
        'DB�s������
        OleDbConnection1.Close()

        '�~
        DYear.Items.Clear()
        For i = 2004 To 2020
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If ListItem1.Value = CStr(DateTime.Now.Year) Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next

        '��
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

        DDescription.Text = ""  '����
        DMakerTime.Text = CStr(DateTime.Now.Today)  '�@����
        BMakerDate.Attributes("onclick") = "calendarPicker('Form1.DMakerTime');"  '������
    End Sub

    Private Sub BSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim Message As String = ""

        'Check�W���ɮ�Size�ή榡
        If ErrCode = 0 Then
            If DDocFileName.Visible Then
                If DDocFileName.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DDocFileName)
                End If
            End If
        End If

        If ErrCode = 0 Then
            SaveData()
            Response.Write(YKK.ShowMessage("�ɮפW�Ǥθ���x�s���\"))
            SetInitData()          '�]�w�e����l��
        Else
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�x�s���ErrCode=0

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �x�s���
    '**
    '*****************************************************************
    Sub SaveData()
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("DocFilePath"))
        Dim FileName As String
        Dim SQL As String
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)    '�W�Ǥ��
        Dim NowDateTime As String = CStr(DateTime.Now.Today) + " " + _
                                    CStr(DateTime.Now.Hour) + ":" + _
                                    CStr(DateTime.Now.Minute) + ":" + _
                                    CStr(DateTime.Now.Second)                '�{�b���


        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQL = "Insert into F_UPDocument "
        SQL = SQL + "( "
        SQL = SQL + "Sts, Class, Type, Year, Month, "                      '1~5
        SQL = SQL + "DocFileName, Description, Maker, MakerTime, "         '6~9
        SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "      '10~13
        SQL = SQL + ")  "
        SQL = SQL + "VALUES( "
        SQL = SQL + " '1', "                                 '���A(0:����,1:�w��NG,2:�w��OK)
        SQL = SQL + " '" + DClass.SelectedValue + "', "      '���O
        SQL = SQL + " '" + DType.SelectedValue + "', "       '�ʽ�
        SQL = SQL + " '" + DYear.SelectedValue + "', "       '�~
        If DType.SelectedValue = "0" Then
            SQL = SQL + " '" + DMonth.SelectedValue + "', "      '��
        Else
            SQL = SQL + " '', "      '��
        End If
        If DDocFileName.Visible Then
            If DDocFileName.PostedFile.FileName <> "" Then   '�W���ɮ�
                FileName = "Doc" & "-" & UploadDateTime & "-" & Right(DDocFileName.PostedFile.FileName, InStr(StrReverse(DDocFileName.PostedFile.FileName), "\") - 1)
                Try
                    DDocFileName.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQL = SQL + " '" + FileName + "', "                 '�W���ɮצW
        SQL = SQL + " '" + DDescription.Text + "', "        '²��
        SQL = SQL + " '" + DMaker.SelectedValue + "', "     '�@����
        SQL = SQL + " '" + DMakerTime.Text + "', "          '�@����
        SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "  '�@����
        SQL = SQL + " '" + NowDateTime + "', "                      '�@���ɶ�
        SQL = SQL + " '" + "" + "', "                               '�ק��
        SQL = SQL + " '" + NowDateTime + "' "                       '�ק�ɶ�
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
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".doc", ".xls"}   '�w�q���\���ɮ׮榡
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '���o�ɮ׮榡
        For i = 0 To allowedExtensions.Length - 1           '�v�@�ˬd���\���榡���O�_���W�Ǫ��榡
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check�W���ɮ�Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If
    End Function

End Class
