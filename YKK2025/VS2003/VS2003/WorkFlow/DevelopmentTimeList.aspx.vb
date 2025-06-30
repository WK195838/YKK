Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class DevelopmentTimeList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DEYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DEMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label

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
    Dim pFormNo As String
    Dim NowYear As String
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim SCreateDate, ECreateDate As String       ''�s�ϩe�U�ѩe�U���(�_,��)


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "DevelopmentTimeList.aspx"

        SetParameter()
        If Not Me.IsPostBack Then
            SetYear()
            SetBuyer()
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
        NowYear = CStr(DateTime.Now.Year)

        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �z��s�ϩe�U�ѩe�U�~��
    '**
    '*****************************************************************
    Sub SetYear()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "SELECT CONVERT(varchar(4), CreateTime, 120) AS SetYear FROM F_MapSheet "
        SQL = SQL + "GROUP BY CONVERT(varchar(4), CreateTime, 120)"

        DSYear.Items.Clear()
        DEYear.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Year")
        DBTable1 = DBDataSet1.Tables("Year")
        '�N�z��X���~���[�JDropDownList
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("SetYear")
            ListItem1.Value = DBTable1.Rows(i).Item("SetYear")
            If ListItem1.Value = NowYear Then ListItem1.Selected = True
            DSYear.Items.Add(ListItem1)
            DEYear.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �z����BUYER
    '**
    '*****************************************************************
    Sub SetBuyer()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()


        If CDate(DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1") > CDate(DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1") Then
            SCreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
            ECreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
        Else
            SCreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
            ECreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
        End If

        SQL = "SELECT a.Buyer FROM F_MapSheet a, T_WaitHandle b "       '�b�e�U������Ӥ�����s�ϩe�U�Ѥ�,���u�{130���z��X��
        SQL = SQL + "WHERE (a.CreateTime > DATEADD(d, 0, '" + SCreateDate + "')) AND (a.CreateTime < DATEADD(m, 1, '" + ECreateDate + "')) "
        SQL = SQL + "AND NOT(a.Sts = '3') AND a.Formno = b.Formno AND a.Formsno = b.Formsno and b.Step = '130' "
        SQL = SQL + "AND NOT(b.CompletedTime IS NULL) GROUP BY a.Buyer"

        DBuyer.Items.Clear()
        DBuyer.Items.Add("ALL")
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "BuyerName")
        DBTable1 = DBDataSet1.Tables("BuyerName")
        '�N�z��X��BUYER�[�JDropDownListr 
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Buyer")
            ListItem1.Value = DBTable1.Rows(i).Item("Buyer")
            DBuyer.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        If DBuyer.SelectedValue <> "" Then
            CreateWorkTable()
        Else
            DataGrid1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Label4.Visible = False
        End If
    End Sub

    Sub CreateWorkTable()
        Dim SQL, SQL1, SQL2, wStep As String
        Dim i, j, Count As Integer
        Dim MapModStsCount As Integer   '�ϭ��ק窱�A��NG�Ψ������ƶq
        Dim MapNo, MapCreateTime, MapFinishTime, MapModTime, MapSts As String               '�ϸ�,ø�ϰ_��ɶ�,ø�ϧ����ɶ�,ø�ϸg�L�ɶ�,ø�϶}�o���A
        Dim MapModCount As Integer      'ø�ϭק隸��
        Dim ManufInCreateTime, ManufInFinishTime, ManufInModTime, ManufInSts, MapToInET As String     '���s�_��ɶ�,���s�����ɶ�,���s�g�L�ɶ�,���s�}�o���A,ø��~���s�g�L�ɶ�
        Dim ManufInModCount As Integer  '���s�ק隸��
        Dim ManufOutCreateTime, ManufOutFinishTime, ManufOutModTime, ManufOutSts, MapToOutET As String '�~�`�_��ɶ�,�~�`�����ɶ�,�~�`�g�L�ɶ�,�~�`�}�o���A,ø��~�~�`�g�L�ɶ�
        Dim ManufOutModCount As Integer '�~�`�ק隸��
        Dim Buyer As String
        Dim MOSUrl, MISUrl As String    '�~�`,���s�s���Ȧs

        Dim WorkTable As String = "Temp_DevelopmentTimeList_" & Request.Cookies("UserID").Value

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBDataSet3 As New DataSet
        Dim DBDataSet4 As New DataSet
        Dim DBDataSet5 As New DataSet
        Dim DBDataSet6 As New DataSet
        Dim DBDataSet7 As New DataSet
        Dim DBDataSet8 As New DataSet
        Dim DBTable1 As DataTable
        Dim DBDataRow As DataRow

        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection

        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        If CDate(DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1") > CDate(DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1") Then
            SCreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
            ECreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
        Else
            SCreateDate = DSYear.SelectedValue + "-" + DSMonth.SelectedValue + "-1"
            ECreateDate = DEYear.SelectedValue + "-" + DEMonth.SelectedValue + "-1"
        End If

        DataGrid1.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True

        If DBuyer.SelectedValue <> "ALL" Then
            SQL1 = "AND a.Buyer = '" + DBuyer.SelectedValue + "' "
        End If

        '�z�����1.�~��,2.BUYER,3.���A�D����,4.Formno & Formsno�ۦP,5.�s�ϩe�U�Ѧ�130�u�{
        SQL = "SELECT a.No,a.Buyer,a.Level,a.Sts,a.CreateTime,b.CompletedTime,a.ModMap,a.MapNo FROM F_MapSheet a,T_WaitHandle b "
        SQL = SQL + "WHERE  (a.CreateTime > DATEADD(d, 0, '" + SCreateDate + "')) AND (a.CreateTime < DATEADD(m, 1, '" + ECreateDate + "')) "
        SQL = SQL + SQL1    'Buyer
        SQL = SQL + "AND NOT (a.Sts = '3') AND a.FormNo = b.FormNo AND a.Formsno = b.Formsno "
        SQL = SQL + "AND b.Step = '130' "
        SQL = SQL + "AND NOT(b.CompletedTime IS NULL) "
        SQL = SQL + "ORDER BY a.Buyer,a.CreateTime"

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "MapSheetInfo")

        'Call Stored Procedure Create WorkTable 
        SQL = "Exec sp_Temp_DevelopmentTimeList '" & WorkTable & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()


        For i = 0 To DBDataSet1.Tables("MapSheetInfo").Rows.Count - 1
            SQL1 = ""
            SQL2 = ""
            MapCreateTime = CStr(DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("createtime"))
            MapModCount = 0
            MapFinishTime = ""
            MapModTime = ""
            MapNo = DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("MapNo")
            MapFinishTime = CStr(DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("CompletedTime"))
            MapSts = DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("Sts")

            If DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("ModMap") = 1 Then    '�P�_�O�_���ϭ��ק�
                MapModStsCount = 0
                '�z�����ϸ��ۦP
                SQL = "SELECT FormNo,Formsno,Sts,MapNo FROM F_MapModSheet WHERE OriMapNo = '" + MapNo + " ' ORDER BY Unique_ID DESC"

                DBDataSet2.Clear()
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet2, "MapModSheetInfo")

                MapModCount = DBDataSet2.Tables("MapModSheetInfo").Rows.Count       '�O��ø�ϭק隸��

                For j = 0 To MapModCount - 1
                    '�z�����1.Formno & Formsno�ۦP,2.�s�ϩe�U�Ѧ�130�u�{
                    SQL = "SELECT CompletedTime FROM T_WaitHandle WHERE FormNo = '"
                    SQL = SQL + DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Formno") + "' AND Formsno = '"
                    SQL = SQL + CStr(DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Formsno")) + "' AND Step = '130' "
                    SQL = SQL + "AND NOT(CompletedTime IS NULL) "

                    DBDataSet3.Clear()
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet3, "WaitHandleInfo")

                    '�P�_�O�_��130�u�{�B���A������OK�ζ}�o��
                    If DBDataSet3.Tables("WaitHandleInfo").Rows.Count >= 1 Then 'And DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Sts") < 2 Then
                        MapSts = DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("Sts")
                        MapFinishTime = CStr(DBDataSet3.Tables("WaitHandleInfo").Rows(0).Item("CompletedTime"))
                        Exit For
                    End If
                Next j

                Count = 0

                '�N�Ҧ��ϭ��ק諸MapNo��X��
                For j = 0 To MapModCount - 1
                    If DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("mapno") <> "" Then    '�P�_�ϸ��O�_���ť�
                        If Count = 0 Then   '�P�_�O�_���Ĥ@��
                            SQL1 = SQL1 + " a.MapNo = '" + DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("mapno") + "' "
                            Count = 1
                        Else
                            SQL1 = SQL1 + "OR a.MapNo = '" + DBDataSet2.Tables("MapModSheetInfo").Rows(j).Item("mapno") + "' "
                        End If
                    End If
                Next j

                If SQL1 <> "" Then      '�P�_SQL1�O�_���ť�
                    SQL2 = "SELECT a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufInSheet a "
                    SQL2 = SQL2 + "LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE "
                    SQL2 = SQL2 + SQL1
                    SQL2 = SQL2 + "GROUP BY a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl "
                    SQL2 = SQL2 + "UNION "
                    SQL1 = "SELECT a.Unique_ID,LEFT(a.No,1) AS No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufOutSheet a LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE " + SQL1
                    SQL1 = SQL1 + "UNION "
                End If
            End If

            If MapFinishTime <> "" Then '�P�_ø�ϧ����ɶ��O�_���ť�,�Y���O����,�h��GetTime�p��g�L�ɶ�
                MapModTime = GetTime(CDate(MapCreateTime), CDate(MapFinishTime))
            End If

            '�H�s�ϩιϭ��ק諸�ϸ���X�~�`�ϸ��O�ۦP��RECORD
            SQL = SQL1 + "SELECT a.Unique_ID,LEFT(a.No,1) AS No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufOutSheet a "
            SQL = SQL + "LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE "
            SQL = SQL + "a.MapNo = '" + MapNo + "' "
            SQL = SQL + "GROUP BY a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl "
            SQL = SQL + "ORDER BY a.Unique_ID DESC"

            DBDataSet4.Clear()
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet4, "ManufOutSheetInfo")

            ManufOutCreateTime = ""
            ManufOutFinishTime = ""
            ManufOutModTime = ""
            ManufOutSts = ""
            MapToOutET = ""
            ManufOutModCount = 0
            MOSUrl = ""

            '�P�_�O�_�����
            If DBDataSet4.Tables("ManufOutSheetInfo").Rows.Count > 0 Then
                wStep = "160"   '�D�ۥD�}�oSTEP160
                ManufOutModCount = DBDataSet4.Tables("ManufOutSheetInfo").Rows.Count    '�~�`���ק隸��
                ManufOutCreateTime = DBDataSet4.Tables("ManufOutSheetInfo").Rows(ManufOutModCount - 1).Item("CreateTime")   '�إߤ�����Ĥ@�����������
                MOSUrl = DBDataSet4.Tables("ManufOutSheetInfo").Rows(0).Item("ViewUrl")
                ManufOutSts = DBDataSet4.Tables("ManufOutSheetInfo").Rows(0).Item("Sts")

                '�P�_�O�_��E�}�Y���e�U��(�ۥD�}�o)
                For j = 0 To ManufOutModCount - 1
                    If DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("No") <> "" And DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("No") = "E" Then
                        wStep = "165"   '�Y���ۥD�}�oSTEP��165
                        Exit For
                    End If
                Next j

                For j = 0 To ManufOutModCount - 1
                    '�z�����1.Formno & Formsno�ۦP,2.�~�`�e�U�Ѧ�160��165�u�{
                    SQL = "SELECT CompletedTime FROM T_WaitHandle WHERE FormNo = '"
                    SQL = SQL + DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("Formno") + "' AND Formsno = '"
                    SQL = SQL + CStr(DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("Formsno")) + "' AND STEP = '"
                    SQL = SQL + wStep + "' "
                    SQL = SQL + "AND NOT(CompletedTime IS NULL) "

                    DBDataSet5.Clear()
                    Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter5.Fill(DBDataSet5, "WaitHandleInfo")

                    '�P�_�O�_��160��165�u�{
                    If DBDataSet5.Tables("WaitHandleInfo").Rows.Count >= 1 Then
                        ManufOutSts = DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("Sts")
                        ManufOutFinishTime = CStr(DBDataSet5.Tables("WaitHandleInfo").Rows(0).Item("CompletedTime"))
                        MOSUrl = DBDataSet4.Tables("ManufOutSheetInfo").Rows(j).Item("ViewUrl")
                        Exit For
                    End If
                Next j
            End If

            '�P�_�~�`�إߤΧ����ɶ��O�_���ť�,�Y���O����,�h��GetTime�p��g�L�ɶ�
            If ManufOutCreateTime <> "" And ManufOutFinishTime <> "" Then
                ManufOutModTime = GetTime(CDate(ManufOutCreateTime), CDate(ManufOutFinishTime))
                MapToOutET = GetTime(CDate(MapCreateTime), CDate(ManufOutFinishTime))
            End If

            '�H�s�ϩιϭ��ק諸�ϸ���X���s�ϸ��O�ۦP��RECORD
            SQL = SQL2 + "SELECT a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl FROM F_ManufInSheet a "
            SQL = SQL + "LEFT OUTER JOIN V_WaitHandle_01 b ON a.FormNo=b.FormNo And a.FormSno=b.FormSno WHERE "
            SQL = SQL + "a.MapNo = '" + MapNo + "' "
            SQL = SQL + "GROUP BY a.Unique_ID,a.No,a.Formno,a.Formsno,a.Sts,a.CreateTime,b.ViewUrl "
            SQL = SQL + "ORDER BY a.Unique_ID DESC"

            DBDataSet6.Clear()
            Dim DBAdapter6 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter6.Fill(DBDataSet6, "ManufInSheetInfo")

            ManufInCreateTime = ""
            ManufInFinishTime = ""
            ManufInModTime = ""
            ManufInSts = ""
            MapToInET = ""
            ManufInModCount = 0
            MISUrl = ""

            '�P�_�O�_�����
            If DBDataSet6.Tables("ManufInSheetInfo").Rows.Count > 0 Then
                wStep = "160"
                ManufInModCount = DBDataSet6.Tables("ManufInSheetInfo").Rows.Count  '���s���ק隸��
                ManufInCreateTime = DBDataSet6.Tables("ManufInSheetInfo").Rows(ManufInModCount - 1).Item("CreateTime")  '�إߤ�����Ĥ@�����������
                MISUrl = DBDataSet6.Tables("ManufInSheetInfo").Rows(0).Item("ViewUrl")
                ManufInSts = DBDataSet6.Tables("ManufInSheetInfo").Rows(0).Item("Sts")

                For j = 0 To ManufInModCount - 1
                    '�z�����1.Formno & Formsno�ۦP,2.���s�e�U�Ѧ�160�u�{
                    SQL = "SELECT CompletedTime FROM T_WaitHandle WHERE FormNo = '"
                    SQL = SQL + DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("Formno") + "' AND Formsno = '"
                    SQL = SQL + CStr(DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("Formsno")) + "' AND STEP = '"
                    SQL = SQL + wStep + "' "
                    SQL = SQL + "AND NOT(CompletedTime IS NULL) "

                    DBDataSet7.Clear()
                    Dim DBAdapter7 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter7.Fill(DBDataSet7, "WaitHandleInfo")

                    '�P�_�O�_��160�u�{
                    If DBDataSet7.Tables("WaitHandleInfo").Rows.Count >= 1 Then
                        ManufInSts = DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("Sts")
                        ManufInFinishTime = CStr(DBDataSet7.Tables("WaitHandleInfo").Rows(0).Item("CompletedTime"))
                        MISUrl = DBDataSet6.Tables("ManufInSheetInfo").Rows(j).Item("ViewUrl")
                        Exit For
                    End If
                Next j
            End If

            '�P�_���s�إߤΧ����ɶ��O�_���ť�,�Y���O����,�h��GetTime�p��g�L�ɶ�
            If ManufInCreateTime <> "" And ManufInFinishTime <> "" Then
                ManufInModTime = GetTime(CDate(ManufInCreateTime), CDate(ManufInFinishTime))
                MapToInET = GetTime(CDate(MapCreateTime), CDate(ManufInFinishTime))
            End If

            '�إ�Table
            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "MNo, Buyer, MLevel, MapCreateTime, MapFinishTime, MapModTime, MapModCount, MapSts, "
            SQL = SQL + "ManufInCreateTime, MISUrl, ManufInFinishTime, ManufInModTime, ManufInModCount, ManufInSts, MapToInET, "
            SQL = SQL + "ManufOutCreateTime, MOSUrl, ManufOutFinishTime, ManufOutModTime, ManufOutModCount, ManufOutSts, MapToOutET, "
            SQL = SQL + "CreateUser, CreateTime "
            SQL = SQL + ") "
            SQL = SQL + "VALUES ("
            SQL = SQL + " '" + DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("No") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("Buyer") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("MapSheetInfo").Rows(i).Item("Level") + "', "
            SQL = SQL + " '" + MapCreateTime + "', "
            SQL = SQL + " '" + MapFinishTime + "', "
            SQL = SQL + " '" + MapModTime + "', "
            SQL = SQL + " '" + CStr(MapModCount) + "', "
            SQL = SQL + " '" + MapSts + "', "
            SQL = SQL + " '" + ManufInCreateTime + "', "
            SQL = SQL + " '" + MISUrl + "', "
            SQL = SQL + " '" + ManufInFinishTime + "', "
            SQL = SQL + " '" + ManufInModTime + "', "
            SQL = SQL + " '" + CStr(ManufInModCount) + "', "
            SQL = SQL + " '" + ManufInSts + "', "
            SQL = SQL + " '" + MapToInET + "', "
            SQL = SQL + " '" + ManufOutCreateTime + "', "
            SQL = SQL + " '" + MOSUrl + "', "
            SQL = SQL + " '" + ManufOutFinishTime + "', "
            SQL = SQL + " '" + ManufOutModTime + "', "
            SQL = SQL + " '" + CStr(ManufOutModCount) + "', "
            SQL = SQL + " '" + ManufOutSts + "', "
            SQL = SQL + " '" + MapToOutET + "', "
            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '�@����
            SQL = SQL + " '" + NowDateTime + "' "       '�@���ɶ�
            SQL = SQL + ") "

            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()

        Next i

        SQL = "SELECT MNo, Buyer, MLevel, MapCreateTime, MapFinishTime, MapModTime, MapModCount , CASE MapSts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' When '3' Then '�}�o����' Else '' End As MapSts, "
        SQL = SQL + "ManufInCreateTime, MISUrl, ManufInFinishTime, ManufInModTime, CASE ManufInModCount WHEN '0' THEN '' ELSE ManufInModCount END AS ManufInModCount, CASE ManufInSts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' When '3' Then '�}�o����' Else '' End As ManufInSts, MapToInET, "
        SQL = SQL + "ManufOutCreateTime, MOSUrl, ManufOutFinishTime, ManufOutModTime, CASE ManufOutModCount WHEN '0' THEN '' ELSE ManufOutModCount END AS ManufOutModCount, CASE ManufOutSts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' When '3' Then '�}�o����' Else '' End As ManufOutSts,  MapToOutET "
        SQL = SQL + "FROM "
        SQL = SQL + WorkTable
        SQL = SQL + " ORDER BY Buyer,MapCreateTime"

        Dim DBAdapter8 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter8.Fill(DBDataSet8, "DevelopmentTimeList")
        DBTable1 = DBDataSet8.Tables("DevelopmentTimeList")

        ''�]�w�Ĥ@�C�βĤG�C���Y���e
        'DBDataRow = DBTable1.NewRow()
        'DBDataRow.Item(0) = "�e�UNo"
        'DBDataRow.Item(1) = "Buyer"
        'DBDataRow.Item(2) = "ø�϶��q"
        'DBDataRow.Item(3) = "���s�e�U��"
        'DBDataRow.Item(4) = "�~�`�e�U��"

        'DBTable1.Rows.InsertAt(DBDataRow, 0)    '���J�b�Ĥ@�C

        'DBDataRow = DBTable1.NewRow()
        'DBDataRow.Item(0) = "�e�UNo"
        'DBDataRow.Item(1) = "Buyer"
        'DBDataRow.Item(2) = "�e�U�_�l��"
        'DBDataRow.Item(3) = "�̲ק�����"
        'DBDataRow.Item(4) = "�g�L�ɶ�"
        'DBDataRow.Item(5) = "�ק隸��"
        'DBDataRow.Item(6) = "���A"
        'DBDataRow.Item(7) = "�e�U�_�l��"
        'DBDataRow.Item(8) = "�̲ק�����"
        'DBDataRow.Item(9) = "�g�L�ɶ�"
        'DBDataRow.Item(10) = "�}�o����"
        'DBDataRow.Item(11) = "���A"
        'DBDataRow.Item(12) = "�e�U�_�l��"
        'DBDataRow.Item(13) = "�̲ק�����"
        'DBDataRow.Item(14) = "�g�L�ɶ�"
        'DBDataRow.Item(15) = "�}�o����"
        'DBDataRow.Item(16) = "���A"
        'DBTable1.Rows.InsertAt(DBDataRow, 0)    '���J�b�ĤG�C

        DataGrid1.DataSource = DBTable1
        DataGrid1.DataBind()

        SetFieldWidth() '�]�w�������e

        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �g�L�ɶ��p��
    '**
    '*****************************************************************
    Function GetTime(ByVal SDate As Date, ByVal EDate As Date) As Integer
        Dim i, TmpHour As Integer
        Dim TmpDate, TmpYMD As Date
        Dim SQL As String
        Dim SSeqNo, ESeqNo As Integer

        Dim DBDataSetA As New DataSet
        Dim DBDataSetB As New DataSet

        Dim OleDBCommand2 As New OleDbCommand
        Dim OleDbConnection2 As New OleDbConnection

        OleDbConnection2.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection2.Open()

        '�@�⦸,�@���}�l,�@������
        For i = 1 To 2
            If i = 1 Then TmpDate = SDate Else TmpDate = EDate
            TmpYMD = TmpDate.Date
            TmpHour = CInt(TmpDate.Hour) * 60 + CInt(TmpDate.Minute)    '�N�ɶ��ഫ����

            SQL = "SELECT MIN(SeqNo) AS SeqNo FROM M_Calendar WHERE YMD = '"
            SQL = SQL + TmpYMD + "' AND Hour >= '"

            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
            'SQL = SQL + CStr(TmpHour) + "' AND Active = '1' AND Depo = 'cl'"
            '
            '�s�զ�ƾ�
            SQL = SQL + CStr(TmpHour) + "' AND Active = '1' AND Depo = 'CL1'"     '���c��ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
            'Modify-End

            DBDataSetA.Clear()
            Dim DBAdapterA As New OleDbDataAdapter(SQL, OleDbConnection2)
            DBAdapterA.Fill(DBDataSetA, "MapModTime")

            '�P�_�O�_�������
            If IsDBNull(DBDataSetA.Tables("MapModTime").Rows(0).Item("SeqNo")) <> True Then
                If i = 1 Then
                    SSeqNo = DBDataSetA.Tables("MapModTime").Rows(0).Item("SeqNo")
                Else
                    ESeqNo = DBDataSetA.Tables("MapModTime").Rows(0).Item("SeqNo")
                End If
            Else
                SQL = "SELECT MIN(SeqNo) AS SeqNo FROM M_Calendar WHERE YMD > '"

                'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                SQL = SQL + TmpYMD + "' AND Active = '1' AND Depo = 'cl'"
                '
                '�s�զ�ƾ�
                SQL = SQL + TmpYMD + "' AND Active = '1' AND Depo = 'CL1'"     '���c��ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
                'Modify-End

                DBDataSetB.Clear()
                Dim DBAdapterB As New OleDbDataAdapter(SQL, OleDbConnection2)
                DBAdapterB.Fill(DBDataSetB, "MapModTime")
                If i = 1 Then
                    SSeqNo = DBDataSetB.Tables("MapModTime").Rows(0).Item("SeqNo")
                Else
                    ESeqNo = DBDataSetB.Tables("MapModTime").Rows(0).Item("SeqNo")
                End If
            End If
        Next i
        OleDbConnection2.Close()
        GetTime = (ESeqNo - SSeqNo) * 10 / 60    '�p��g�L�ɶ�(�p��)
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�������e
    '**
    '*****************************************************************
    Sub SetFieldWidth()
        Dim i As Integer
        Dim TempWidth As Integer

        'DataGrid1.Items(0).Cells(0).RowSpan = 2     '�N�Ĥ@�C�Ĥ@����ܬ����V����
        'DataGrid1.Items(0).Cells(1).RowSpan = 2     '�N�Ĥ@�C�Ĥ@����ܬ����V����
        'DataGrid1.Items(0).Cells(2).ColumnSpan = 5  '�N�Ĥ@�C�ĤG����ܬ���V����
        'DataGrid1.Items(0).Cells(3).ColumnSpan = 5  '�N�Ĥ@�C�ĤT����ܬ���V����
        'DataGrid1.Items(0).Cells(4).ColumnSpan = 5  '�N�Ĥ@�C�ĥ|����ܬ���V����

        'For i = 5 To DataGrid1.Columns.Count - 1
        '    DataGrid1.Items(0).Cells(i).Visible = False '���þ�V����᭱12��
        'Next i

        'DataGrid1.Items(1).Cells(0).Visible = False '���ê��V����᭱1��
        'DataGrid1.Items(1).Cells(1).Visible = False '���ê��V����᭱1��

        ''�]�w�Ĥ@�C�ĤG�B�T�B�|��e�׬�
        'DataGrid1.Items(0).Cells(2).Width = New Unit(240)
        'DataGrid1.Items(0).Cells(3).Width = New Unit(240)
        'DataGrid1.Items(0).Cells(4).Width = New Unit(240)

        '�N�Ĥ@�B�G�C�Ҧ����r�]�w���m��
        'For i = 0 To DataGrid1.Columns.Count - 1
        '    DataGrid1.Items(0).Cells(i).HorizontalAlign = HorizontalAlign.Center
        'DataGrid1.Items(1).Cells(i).HorizontalAlign = HorizontalAlign.Center
        'Next i

        '�]�w�Ĥ@�B�G�C�Ҧ����I���Φr���C��
        'DataGrid1.Items(0).BackColor = Color.FromArgb(13395456)
        'DataGrid1.Items(0).ForeColor = System.Drawing.Color.White
        'DataGrid1.Items(1).BackColor = Color.FromArgb(13395456)
        'DataGrid1.Items(1).ForeColor = System.Drawing.Color.White



        For i = 0 To 19
            Select Case i
                Case 0
                    TempWidth = 80
                Case 1
                    TempWidth = 80
                Case 2
                    TempWidth = 40
                Case 3
                    TempWidth = 80
                Case 4
                    TempWidth = 80
                Case 5
                    TempWidth = 30
                Case 6
                    TempWidth = 20
                Case 7
                    TempWidth = 30
                Case 8
                    TempWidth = 80
                Case 9
                    TempWidth = 80
                Case 10
                    TempWidth = 30
                Case 11
                    TempWidth = 20
                Case 12
                    TempWidth = 30
                Case 13
                    TempWidth = 30
                Case 14
                    TempWidth = 80
                Case 15
                    TempWidth = 80
                Case 16
                    TempWidth = 30
                Case 17
                    TempWidth = 20
                Case 18
                    TempWidth = 30
                Case 19
                    TempWidth = 30
            End Select
            DataGrid1.Columns.Item(i).ItemStyle.Width = New Unit(TempWidth)
            DataGrid1.Columns.Item(i).HeaderStyle.Width = New Unit(TempWidth)
            DataGrid1.Width = New Unit(1161)
            'width = 80+80+40+80+80+30+20+40+80+80+30+20+40+30+80+80+30+20+40+30=950,��u=21(20+1),�w�d�Ŷ�4*2(����)*20=160,total=1161
        Next
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��Excel�@�ε{��
    '**
    '*****************************************************************
    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        ''pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '�{���O���P
        CreateWorkTable()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=Commission_List.xls")     '�{���O���P
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '�{���O���P
    End Sub

    Private Sub DSYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DSYear.SelectedIndexChanged
        SetBuyer()
    End Sub

    Private Sub DSMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DSMonth.SelectedIndexChanged
        SetBuyer()
    End Sub

    Private Sub DEYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DEYear.SelectedIndexChanged
        SetBuyer()
    End Sub

    Private Sub DEMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DEMonth.SelectedIndexChanged
        SetBuyer()
    End Sub
End Class
