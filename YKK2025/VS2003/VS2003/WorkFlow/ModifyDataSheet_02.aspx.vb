Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ModifyDataSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DModifyDataSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DIPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFDateTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DIContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMContent2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMContent1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSheet As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMReasonType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DComNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DStatus As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LSheet1 As System.Web.UI.WebControls.HyperLink

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
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim NowDateTime As String       '�{�b����ɶ�

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "ModifyDataSheet_02.aspx"

        SetParameter()          '�]�w�@�ΰѼ�

        If Not Me.IsPostBack Then   '���OPostBack
            ShowFormData()      '��ܪ����
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
    End Sub


    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ModifyDataSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModifyDataSheet")
        If DBDataSet1.Tables("F_ModifyDataSheet").Rows.Count > 0 Then

            DNo.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("No")                         'No
            DDate.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Date")                     '���
            SetFieldData("Division", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Division"))    '�e�U����
            SetFieldData("Person", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Person"))        '�e�U���
            SetFieldData("MDivision", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MDivision"))  '�קﳡ��
            SetFieldData("MPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MPerson"))      '�ק���
            SetFieldData("Sheet", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Sheet"))          '�e�U��
            DComNo.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("ComNo")                   '�e�UNo
            SetFieldData("Status", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Status"))        '�e�U��
            SetFieldData("WDivision", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("WDivision"))  '�ݳB�z�u�{����
            SetFieldData("WPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("WPerson"))      '�ݳB�z�u�{���

            SetFieldData("MReasonType", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MReasonType"))   '�ק�z�����O
            DMReason.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MReason")               '�ק�z��
            DMContent1.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MContent1")           '�ק鷺�e-1
            DMContent2.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MContent2")           '�ק鷺�e-2

            DIContent.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("IContent")             '��ڭק鷺�e
            DFDateTime.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("FDateTime")           '�w�w�����ɶ�
            SetFieldData("IPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("IPerson"))      '��ڭק���

        End If
        '���s��

        SQL = " select  formno,tablename1  from M_form"
        SQL &= " where formname = '" + Trim(DSheet.SelectedValue) + "'"
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        Dim ds As New Data.DataSet
        DBAdapter3.Fill(ds)
        Dim SheetFormNo As String
        Dim SheetFormSno As String

        Dim dt As DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            SheetFormNo = dt.Rows(0)("formno")
            SQL = " select formsno from f_" + dt.Rows(0)("tablename1")
            SQL &= " where no='" + Trim(DComNo.Text) + "'"
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            Dim ds1 As New Data.DataSet
            DBAdapter4.Fill(ds1)
            Dim dt1 As DataTable = ds1.Tables(0)
            Dim sUrl As String
            Dim cun As Integer

            If dt1.Rows.Count > 0 Then
                SheetFormSno = dt1.Rows(0)("formsno")
                sUrl = dt.Rows(0)("tablename1") + "_02.aspx?pFormNo=" & SheetFormNo & "&pFormSno=" & SheetFormSno
                'Dim sUrlStr = Mid(sUrl, 11, sUrl.Length - 1)
                LSheet1.NavigateUrl = sUrl

            End If
        End If

        DFormSno.Text = "�渹�G" & CStr(wFormSno)
        'DB�s������
        OleDbConnection1.Close()
    End Sub


    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        idx = 0

        '�e�U����
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.Cookies("UserID").Value & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '�e�U���
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.Cookies("UserID").Value & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '�קﳡ��
        If pFieldName = "MDivision" Then
            DMDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where Active = '1' "
                SQL = SQL & " Group By DivName "
                SQL = SQL & " Order By DivName "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = DDivision.SelectedValue Then ListItem1.Selected = True
                    DMDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '�ק���
        If pFieldName = "MPerson" Then
            DMPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where DivName = '" & DMDivision.SelectedValue & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = DPerson.SelectedValue Then ListItem1.Selected = True
                    DMPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '�ݳB�z�u�{����
        If pFieldName = "WDivision" Then
            DWDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where Active = '1' "
                SQL = SQL & " Group By DivName "
                SQL = SQL & " Order By DivName "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    If pName <> "ZZZZZZ" Then
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Value = DBTable1.Rows(i).Item("DivName")
                        If ListItem2.Value = pName Then ListItem2.Selected = True
                        DWDivision.Items.Add(ListItem2)
                    Else
                        If i = 0 Then
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = ""
                            ListItem1.Value = ""
                            ListItem1.Selected = True
                            DWDivision.Items.Add(ListItem1)
                        End If
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Value = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Selected = False
                        DWDivision.Items.Add(ListItem2)
                    End If
                Next
            End If
        End If

        '�ݳB�z�u�{���
        If pFieldName = "WPerson" Then
            DWPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")

                For i = 0 To DBTable1.Rows.Count - 1
                    If pName <> "ZZZZZZ" Then
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        If ListItem2.Value = pName Then ListItem2.Selected = True
                        DWPerson.Items.Add(ListItem2)
                    Else
                        If i = 0 Then
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = ""
                            ListItem1.Value = ""
                            ListItem1.Selected = True
                            DWPerson.Items.Add(ListItem1)
                        End If
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Selected = False
                        DWPerson.Items.Add(ListItem2)
                    End If
                Next
            End If
        End If

        '�e�U��
        If pFieldName = "Sheet" Then
            DSheet.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSheet.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select FormName From M_Form "
                SQL = SQL & " Where FormNo < '800000' "
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Form")
                DBTable1 = DBDataSet1.Tables("M_Form")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("FormName")
                    ListItem1.Value = DBTable1.Rows(i).Item("FormName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSheet.Items.Add(ListItem1)
                Next
            End If
        End If

        '���A
        If pFieldName = "Status" Then
            DStatus.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DStatus.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='STATUS' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DStatus.Items.Add(ListItem1)
                Next
            End If
        End If

        '�ק�z�����O
        If pFieldName = "MReasonType" Then
            DMReasonType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMReasonType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='MODIFYREASON' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMReasonType.Items.Add(ListItem1)
                Next
            End If
        End If

        '��ڭק���
        If pFieldName = "IPerson" Then
            DIPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DIPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='MODIFYPERSON' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DIPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        'DB�s������
        OleDbConnection1.Close()
    End Sub

End Class
