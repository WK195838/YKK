Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls
Imports System.Drawing
Imports System.Windows.forms
 


Partial Class DASW_DISPOSALUpload01
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '�U���
    Dim Attribute(60) As Integer    '�U����ݩ�    
    Dim Top As String               '�w�]�������m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wbFormSno As Integer        '�s��_��-���y����
    Dim wbStep As Integer           '�s��_��-�u�{�N�X
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID

    Dim wAgentID As String          '�Q�N�z�HID
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim wNextGate As String         '�U�@���H
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '�x�_��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End

    Dim wUserName As String = ""    '�m�W�N�z��
    Dim HolidayList As New List(Of Integer) '�ΥH�O�����骺�����ޭ�
    Dim indexList As New List(Of Integer)   '�ΥH�O�����ݩ�������������ޭ�
    Dim DateList As New List(Of String)     '�ΥH�O���ҿ�����@�P���

    ''' <summary>
    ''' �H�U���@�Ψ禡���ŧi
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim isOK As Boolean = True
    Dim Message As String = ""
    Dim OpenDir2 As String = ""
    Dim SignDate, OpenDate As String
    Dim UploadName As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '���OPostBack
            SetParameter()          '�]�w�@�ΰѼ�
            SetFieldAttribute("")
        End If


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '�{�b���
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        Dim SQL As String


        '���̤γ��� 

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DAPPDate.Text = Now.ToString("yyyy/MM/dd") '�{�b���
        DAPPDepo.Text = DBUser.Rows(0).Item("Divname")
        DAppName.Text = DBUser.Rows(0).Item("Username")
        DSMONTH.Text = Mid(DAPPDate.Text, 1, 7)


        ' Dim i As Integer
        'SQL = " select max(disposalym)disposalym from F_DISPOSALSheet order by disposalym desc "
        'Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        ' DSMONTH.Items.Clear()
        '        DSMONTH.Items.Add("")
        'For i = 0 To DBAdapter1.Rows.Count - 1
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = DBAdapter1.Rows(i).Item("disposalym")
        'ListItem1.Value = DBAdapter1.Rows(i).Item("disposalym")
        '   DSMONTH.Items.Add(ListItem1)
        'Next
        ' DSMONTH.Text = Now.ToString("yyyy/MM") '�{�b���
        'DBAdapter1.Clear()

    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     ��������ܤ�����J�ˬd
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '�ظm�����ݩʰ}�C
        '    oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '��渹�X,�u�{���d���X,���W,����ݩ�

        ' SetFieldAttribute(pPost)     '���U����ݩʤ�����J�ˬd���]�w
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     ���U����ݩʤ�����J�ˬd���]�w
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '���U����ݩʤ�����J�ˬd���]�w
        DAPPDate.BackColor = Color.LightGray
        DAPPDate.ReadOnly = True
        DAPPDate.Visible = True
        DAPPDepo.BackColor = Color.LightGray
        DAPPDepo.ReadOnly = True
        DAPPDepo.Visible = True
        DAppName.BackColor = Color.LightGray
        DAppName.ReadOnly = True
        DAppName.Visible = True
        DSMONTH.BackColor = Color.LightGray
        DSMONTH.Visible = True
        DSMONTH.ReadOnly = True

        DFileUpload1.Visible = True
        DFileUpload1.Style.Add("BACKGROUND-COLOR", "GreenYellow")


      




    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     �إߪ��ݿ�J���
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        'rqdVal.Display = ValidatorDisplay.None              ' �]�b�����W�[�JValidationSummary , �G���ұ���Τ@���
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '�W�Ǹ���ˬd����ܰT��
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        If Message <> "" Then
            Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub



    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".xls", ".xlsx"}   '�w�q���\���ɮ׮榡
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '���o�ɮ׮榡
        For i = 0 To allowedExtensions.Length - 1           '�v�@�ˬd���\���榡���O�_���W�Ǫ��榡
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9020
            End If
        Next
        'Check�W���ɮ�Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9030
            End If
        End If
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �P�_�O�_�i�~�����(���Ҹ��)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim Message As String = ""

        If DSMONTH.Text = "" Then
            isOK = False
            ErrCode = 9010

        End If

        If ErrCode = 0 Then
            If DFileUpload1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                ErrCode = UPFileIsNormal(DFileUpload1)
            Else
                isOK = False
                ErrCode = 9040
            End If
        End If


        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "�п�J���"
            If ErrCode = 9020 Then Message = "�W���ɮ׮榡���~9020,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮ�Size���~,�нT�{!"
            If ErrCode = 9040 Then Message = "�W���ɮ׮榡���~9040,�нT�{!"
            uJavaScript.PopMsg(Me, Message)
            GridView1.DataSource = Nothing
            GridView1.DataBind()
            GridView2.DataSource = Nothing
            GridView2.DataBind()
        Else
            isOK = True
        End If

        Return isOK

    End Function

  
    Sub upload()
        Try
            If DFileUpload1.HasFile Then

                '�W�Ǫ���
                Dim FileName1 As String
                UploadName = DFileUpload1.PostedFile.FileName

                Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
                Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                System.Configuration.ConfigurationManager.AppSettings("DISPOSALData")

                '20170912 �N�ɦW�ק令���t��l�ɦW
                Dim fileExtension As String  '���ɦW
                fileExtension = IO.Path.GetExtension(UploadName).ToLower   '���o�ɮ׮榡
                FileName1 = Path1 + CStr(DNo.Text) + UploadDateTime + fileExtension
                DFileUpload1.PostedFile.SaveAs(FileName1)

                '�i�}
                Dim FileName As String = Path.GetFileName(DFileUpload1.PostedFile.FileName)
                Dim Extension As String = Path.GetExtension(DFileUpload1.PostedFile.FileName)
                '  Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

                ' FileName1 = "http://10.245.1.6/DASW/Document/006002/" + CStr(DNo.Text) + UploadDateTime + fileExtension
                'Dim FilePath As String = Server.MapPath(FolderPath + FileName)
                ' DFileUpload1.SaveAs(FilePath)
                'Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Text)
                Import_To_Grid(FileName1, Extension, rbHDR.SelectedItem.Text)




            End If
        Catch ex As Exception
            uJavaScript.PopMsg(Me, "�W���ɮ׮榡���~upload,�нT�{!")

        End Try

    End Sub

    Private Sub Import_To_Grid(ByVal FilePath As String, ByVal Extension As String, ByVal isHDR As String)
        Dim conStr As String = ""
        Select Case Extension
            Case ".xls"
                'Excel 97-03
                conStr = ConfigurationManager.ConnectionStrings("Excel03ConString").ConnectionString
                Exit Select
            Case ".xlsx"
                'Excel 07
                conStr = ConfigurationManager.ConnectionStrings("Excel07ConString").ConnectionString
                Exit Select
        End Select
        conStr = String.Format(conStr, FilePath, isHDR)

        Dim connExcel As New OleDbConnection(conStr)
        Dim cmdExcel As New OleDbCommand()
        Dim oda As New OleDbDataAdapter()
        Dim dt As New DataTable()

        cmdExcel.Connection = connExcel

        'Get the name of First Sheet
        connExcel.Open()
        Dim dtExcelSchema As DataTable
        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        Dim i As Integer
        Dim SheetName As String = ""

        For i = 0 To dtExcelSchema.Rows.Count - 1
            If dtExcelSchema.Rows.Count > 1 Then
                SheetName = dtExcelSchema.Rows(1)("TABLE_NAME").ToString()
                If SheetName = "�u�@��1$" Or SheetName = "���o����$" Then
                Else
                    SheetName = "�u�@��1$"
                End If
            Else
                SheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
            End If

        Next



        connExcel.Close()

        


        'Read Data from First Sheet
        connExcel.Open()
        cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
        oda.SelectCommand = cmdExcel
        oda.Fill(dt)
        connExcel.Close()

        'Bind Data to GridView
        '��ܥ�
        GridView1.Caption = Path.GetFileName(FilePath)
        GridView1.DataSource = dt
        GridView1.DataBind()
        '�פJ��
        GridView2.Caption = Path.GetFileName(FilePath)
        GridView2.DataSource = dt
        GridView2.DataBind()
    End Sub

    Protected Sub PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")
        Dim FileName As String = GridView1.Caption
        Dim Extension As String = Path.GetExtension(FileName)
        Dim FilePath As String = Server.MapPath(FolderPath + FileName)

        Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Text)
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataBind()
    End Sub

    Protected Sub DUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DUpload.Click
        If InputDataOK(0) Then
            upload()
        Else

        End If


    End Sub

    Protected Sub DInsert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DInsert.Click
        '�ˬd�O�_���P����P�ӽФH���ɮסA�Y���h�N���e�����R��
        Dim a As String = ""
     

        Dim uploadflag As Integer = 1
        Dim nullflag As Integer = 1
        Dim SQL As String
        SQL = " select  distinct  appname,smonth   from f_disposaldata"
        SQL = SQL + "  where appname = '" + DAppName.Text + "'"
        SQL = SQL + " and No = '" + DNO1.Text + "'"
        Dim dt1 As DataTable = uDataBase.GetDataTable(sql)
        If dt1.Rows.Count > 0 Then
            Dim MsgBoxFlag As Integer
            MsgBoxFlag = MessageBox.Show(DSMONTH.Text + "�����Ƥw�W�ǹL�A�O�_�n�N��ƲM�ŦA���s�W��?", "AlanStudio", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            If MsgBoxFlag = 1 Then
               
                SQL = " delete  from f_disposaldata"
                SQL = SQL + "  where appname = '" + DAppName.Text + "'"
                SQL = SQL + " and no= '" + DNO1.Text + "'"
                uDataBase.ExecuteNonQuery(SQL)
                uploadflag = 1
            Else
                uploadflag = 2
            End If

        End If
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID

        If uploadflag = 1 Then
            Try

                '�W�Ǩ��Ʈw
                Dim j As Integer
                Dim jSQL As String
                jSQL = ""
                DNo.Text = SetNo(0)
                Dim i As Integer


                For i = 0 To Me.GridView2.Rows.Count - 1 Step i + 1



                    SQL = "Insert into F_DISPOSALData (APPDATE,APPNAME,APPDEPO,SMONTH,No,Seqno,Code,ITEMNAME1,ITEMNAME2,LENGTH,U,Color,LOCATION,ACTUAL,FREE,SIZE,CHAIN,CLS,UNIT,UNITWEIGHT,"
                    SQL = SQL + "WEIGHTKG,DisposalRule,PNStock,COSTA,COSTB,ACTUALAMOUNT,UT2,LASTIN,LASTOUT,"
                    SQL = SQL + "DEPONAME,SALES,DISPOSALREASON,ONEYEAR,CREATEDATE,CREATEUSER)"
                    SQL = SQL + " values('" + DAPPDate.Text + "','" + DAppName.Text + "','" + DAPPDepo.Text + "','" + DSMONTH.Text + "','" + DNO1.Text + "'," + CStr(i + 1) + ","
                    For j = 0 To 26
                      

                        If j = 0 Then
                            If GridView2.Rows(i).Cells(j).Text = "&nbsp;" Then
                                jSQL = "''"
                            Else
                                jSQL = "'" + GridView2.Rows(i).Cells(j).Text + "'"

                            End If

                        Else

                            If GridView2.Rows(i).Cells(j).Text = "&nbsp;" Then
                                jSQL = jSQL + "," + "''"
                            Else
                                jSQL = jSQL + ",'" + GridView2.Rows(i).Cells(j).Text + "'"

                            End If


                        End If

                      
                        a = GridView1.Rows(i).Cells(0).Text
                        If a = "&nbsp;" Then  '�ˬd�Ĥ@��O���ONULL
                            nullflag = 0
                        End If

                    Next

                    SQL = SQL + jSQL + ","
                    SQL = SQL + "getdate(),'" + wApplyID + "')"

                    'If CStr(i + 1) = "51" Then
                    'uJavaScript.PopMsg(Me, a)
                    'End If


                    If nullflag = 1 Then '�Ĥ@��O�ťմN���n�פJ
                        uDataBase.ExecuteNonQuery(SQL)

                    End If


                Next


                uJavaScript.PopMsg(Me, "�W�Ǧ��\")
                ' Send1(wApplyID, "pc037", wApplyID, "006002", "1", "1", "END")
                '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O

                DInsert.Enabled = False
            Catch ex As Exception
                uJavaScript.PopMsg(Me, "�W���ɮ׮榡���~Insert,�нT�{!")

            End Try

        End If


        GridView1.DataSource = Nothing
        GridView1.DataBind()
        GridView2.DataSource = Nothing
        GridView2.DataBind()





    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �s�s�e�UNo 
    '**
    '*****************************************************************
    Function SetNo(ByVal Seq As Integer) As String
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim i As Integer

        'Set�����
        Str2 = CStr(DateTime.Now.Month)  '��
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '��
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = CStr(DateTime.Now.Year) + Str2
        '�~
        'Set�渹
        '�������Ʀ��X��  20150414 Modify by Jessica
        Dim sql As String
        sql = " select  NO,isnull(right(max(no),2),0) as  cun from  F_DISPOSALData  "
        sql = sql + " GROUP BY NO "
        Dim dt1 As DataTable = uDataBase.GetDataTable(sql)
        If dt1.Rows.Count > 0 Then
            Str1 = CStr(CInt(dt1.Rows(0).Item("cun")) + 1)
        Else
            Str1 = "001"
        End If

        For i = Str1.Length To 3 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

       

        If e.Row.RowType = DataControlRowType.Header Then


            '�ˬd���W��

            ' s1 = Trim(e.Row.Cells(26).Text.ToUpper)
            DInsert.Enabled = True

            'Dim a As String
            'a = Trim(e.Row.Cells(0).Text.ToUpper)
            'a = Trim(e.Row.Cells(1).Text.ToUpper)

            If Trim(e.Row.Cells(0).Text.ToUpper) <> "CODE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(1).Text.ToUpper) <> "ITEM NAME 1" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(2).Text.ToUpper) <> "ITEM NAME 2" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(3).Text.ToUpper) <> "LENGTH" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(4).Text.ToUpper) <> "U" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(5).Text.ToUpper) <> "COLOR" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(6).Text.ToUpper) <> "LOCATION" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(7).Text.ToUpper) <> "ACTUAL" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(8).Text.ToUpper) <> "FREE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(9).Text.ToUpper) <> "SIZE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(10).Text.ToUpper) <> "CHAIN" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(11).Text.ToUpper) <> "CLS" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(12).Text.ToUpper) <> "UNIT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(13).Text.ToUpper) <> "UNIT _WEIGHT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(14).Text.ToUpper) <> "WEIGHT(KG)" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(15).Text.ToUpper) <> "���o�ǫh" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(16).Text.ToUpper) <> "PN �ܦ�" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(17).Text.ToUpper) <> "COST A" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(18).Text.ToUpper) <> "COST B" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(19).Text.ToUpper) <> "ACTUAL_AMOUNT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(20).Text.ToUpper) <> "UT2" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(21).Text.ToUpper) <> "LAST IN" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(22).Text.ToUpper) <> "LAST OUT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(23).Text.ToUpper) <> "����" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(24).Text.ToUpper) <> "��~���" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(25).Text.ToUpper) <> "��]" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(26).Text.ToUpper) <> "���~_�ϥζq" Then
                DInsert.Enabled = False
            End If

            'End If




            If DInsert.Enabled = False Then
                Message = "�W�Ǯ榡���~gridview,�нT�{!"
                uJavaScript.PopMsg(Me, Message)

                GridView1.DataSource = Nothing
                GridView1.DataBind()
                GridView2.DataSource = Nothing
                GridView2.DataBind()
            End If

        Else

            '2019/2/14 �W�[ �T����줣���\�ť�
            If Trim(e.Row.Cells(0).Text.ToUpper) <> "&NBSP;" Then
                If Trim(e.Row.Cells(15).Text.ToUpper) = "&NBSP;" Then
                    DInsert.Enabled = False
                End If

                If Trim(e.Row.Cells(23).Text.ToUpper) = "&NBSP;" Then
                    DInsert.Enabled = False
                End If

            End If


          
            If DInsert.Enabled = False Then
                Message = "CODE �� ���o�ǫh �� ���� �����\�ť�!"
                uJavaScript.PopMsg(Me, Message)

                GridView1.DataSource = Nothing
                GridView1.DataBind()
                GridView2.DataSource = Nothing
                GridView2.DataBind()
            End If

        End If



    End Sub

    'SendMail-End
    '**************************************************************************************
    '** ���Ͷl���ƦܧǦC��(Send)
    '**************************************************************************************
    'Send-Start
    Public Function Send1(ByVal pFrom As String, _
                         ByVal pTo As String, _
                         ByVal pApplyID As String, _
                         ByVal pFormNo As String, _
                         ByVal pFormSno As Integer, _
                         ByVal pStep As Integer, _
                         ByVal pType As String) As Integer
        Dim RtnCode As Integer = 0
        Dim NowDateTime As String = Now.ToString("yyyy/MM/dd HH:mm:ss")
        Dim SQL As String
        Dim xFlowType As String = ""
        Dim xStepname As String = ""
        Dim xNo As String = ""
        Dim xFormName As String = ""
        Dim xFromMail As String = ""
        Dim xFromName As String = ""
        Dim xToMail As String = ""
        Dim xToName As String = ""
        Dim xApplyName As String = ""
        Dim xCCMail As String = ""
        Try
        


            Dim dt_Mail As DataTable
            '���W��
            SQL = "SELECT FormName FROM M_Form "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xFormName = dt_Mail.Rows(0).Item("FormName").ToString
            End If
            '�H��̶l��b��,�m�W
            SQL = "Select Mail, UserName From M_Users "
            SQL &= "Where UserID  = '" & pFrom & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xFromMail = dt_Mail.Rows(0).Item("Mail").ToString
                xFromName = dt_Mail.Rows(0).Item("UserName").ToString
            End If
            '����̶l��b��,�m�W
            SQL = "Select Mail, UserName From M_Users "
            SQL &= "Where UserID  = '" & pTo & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xToMail = dt_Mail.Rows(0).Item("Mail").ToString
                xToName = dt_Mail.Rows(0).Item("UserName").ToString
            End If
            '�ӽЪ̩m�W
            xApplyName = dt_Mail.Rows(0).Item("UserName").ToString
            xStepname = DSMONTH.Text + "-" + DAPPDepo.Text + "-" + xFromName
          
            '���Ͷl���ƦܧǦC��
            SQL = "Insert into Q_WaitSend "
            SQL &= "( "
            SQL &= "Sts, FromID, FromMail, FromName, ToID, "
            SQL &= "ToMail, ToName, CCMail, "
            SQL &= "FormNo, FormSno, FormName, Step, StepName, "
            SQL &= "ApplyID, ApplyName, MSG, MSGName, No, "
            SQL &= "CreateTime "
            SQL &= ") "
            SQL &= "VALUES( "
            SQL &= "'0' ,"                              '����A(0:����,1:�ݳB�z)
            SQL &= "'" & pFrom & "' ,"                  '�H���ID
            SQL &= "'" & xFromMail & "' ,"              '�H��̶l��b��
            SQL &= "'" & xFromName & "' ,"              '�H��̩m�W
            SQL &= "'" & pTo & "' ,"                    '�����ID
            SQL &= "'" & xToMail & "' ,"                '����̶l��b��
            SQL &= "'" & xToName & "' ,"                '����̩m�W
            SQL &= "'" & xCCMail & "' ,"                'cc�̶l��b��
            SQL &= "'" & pFormNo & "' ,"                '���No
            SQL &= "'" & CStr(pFormSno) & "' ,"         '���Ǹ�
            SQL &= "'" & xFormName & "' ,"              '���W��
            SQL &= "'" & CStr(pStep) & "' ,"            '�u�{No
            SQL &= "'" & xStepname & "' ,"              '�u�{�W��
            SQL &= "'" & pApplyID & "' ,"               '�ӽЪ�ID
            SQL &= "'" & xApplyName & "' ,"             '�ӽЪ̩m�W
            SQL &= "'" & pType & "' ,"                  '�T�����O�N�X
            SQL &= "'" & xFlowType & "' ,"              '�T�����O�W��
            SQL &= "'" & xNo & "' ,"                    '�e�UNo
            SQL &= "'" & NowDateTime & "' "             '�s�@�ɶ�
            SQL &= ") "
            uDataBase.ExecuteNonQuery(SQL)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function

   
End Class

