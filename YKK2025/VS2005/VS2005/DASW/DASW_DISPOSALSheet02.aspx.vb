Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO

 


Partial Class DASW_DISPOSALSheet02
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
    Dim wUserID As String          '�ϥΪ�ID
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





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�
  
        ShowSheetField("New")   '��������ܤ�����J�ˬd


        ShowFormData()      '��ܪ����

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()

        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        LFormSno.Text = "�渹:" + CStr(wFormSno)
        ' DDISPOSALFILE2.Attributes.Add("onclick", "window.open('file://10.245.1.18/MIS/DASW/D201507001','_blank');return false;")

        '�ˬd�O�_�w�g�L�ꪫ�o�󧹦�
        Dim SQL As String
        SQL = " select * from t_waithandle "
        SQL = SQL + " where formno ='006001'"
        SQL = SQL + " and formsno = '" + CStr(wFormSno) + "'"
        SQL = SQL + " and step  = 140"
        SQL = SQL + " and sts =1"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count = 0 Then
            DDISPOSALFILE2.Visible = False
        End If

      
    End Sub


    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()



        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                             System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")  'WIS-TempPath

        Dim SQL As String
        '�D�ɸ��
        SQL = "Select * From F_DISPOSALSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        '�]�w�ɮ׸��|

        DMKTSIGN.Checked = False
        DCUSTOMERTOLL.Checked = False

        If DBAdapter1.Rows(0).Item("MKTSIGN") = 1 Then
            DMKTSIGN.Checked = True
        End If

        If DBAdapter1.Rows(0).Item("CUSTOMERTOLL") = 1 Then
            DCUSTOMERTOLL.Checked = True
        End If

        DNo.Text = DBAdapter1.Rows(0).Item("No")

        '�}�ҳ��o�����

        '�D�ɸ��
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '6001'"
        SQL = SQL + " and dkey ='DisposalFilePath'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        Dim OpenDir As String = ""
        If DBAdapter3.Rows.Count > 0 Then
            OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text
        End If

        'Dim OpenDir As String
        'OpenDir = "file://10.245.1.18/MIS/DASW/" + DNo.Text

        DDISPOSALFILE2.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")

        DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")
        DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")
        DAppName.Text = DBAdapter1.Rows(0).Item("AppName")
        DDISPOSALREASON.Text = DBAdapter1.Rows(0).Item("DISPOSALREASON")
        ' SetFieldData("DISPOSALREASON", DBAdapter1.Rows(0).Item("DISPOSALREASON"))
        DDISPOSALREASON1.Text = DBAdapter1.Rows(0).Item("DISPOSALREASON1")
        DDUTYDEPO.Text = DBAdapter1.Rows(0).Item("DUTYDEPO")
        'SetFieldData("DUTYDEPO", DBAdapter1.Rows(0).Item("DUTYDEPO"))
        DSales.Text = DBAdapter1.Rows(0).Item("Sales")
        DPLACE.Text = DBAdapter1.Rows(0).Item("PLACE")
        DSTYPE.Text = DBAdapter1.Rows(0).Item("STYPE")

        'SetFieldData("PLACE", DBAdapter1.Rows(0).Item("PLACE"))
        DDISPOSALRULE.Text = DBAdapter1.Rows(0).Item("DISPOSALRULE")
        ' SetFieldData("DISPOSALRULE", DBAdapter1.Rows(0).Item("DISPOSALRULE"))
        DDISPOSALTYPE.Text = DBAdapter1.Rows(0).Item("DISPOSALTYPE")

        DCHINESEREASON.Text = DBAdapter1.Rows(0).Item("CHINESEREASON")
        DJAPANREASON.Text = DBAdapter1.Rows(0).Item("JAPANREASON")
        If DBAdapter1.Rows(0).Item("DISPOSALFILE1") <> "" Then
            LDISPOSALFILE.NavigateUrl = Path & DBAdapter1.Rows(0).Item("DISPOSALFILE1") '���o����
            LDISPOSALFILE.Visible = True
        Else
            LDISPOSALFILE.Visible = False
        End If


        If DBAdapter1.Rows(0).Item("PCNAME") = "" Then
            SQL = " Select  a.*,userid  from M_referp a,M_users b"
            SQL = SQL & "   where  cat = '6001'"
            SQL = SQL & "  and dkey = 'PCNAME'"
            SQL = SQL & " and a.data =b.username "
            SQL = SQL & " AND b.userid ='" + DBAdapter1.Rows(0).Item("PCNAME") + "'"
            SQL = SQL & " order by a.unique_id"
            Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
            If dtReferp.Rows.Count > 0 Then
                If dtReferp.Rows(0).Item("Data") <> "" Then
                    DPCNAME.Text = dtReferp.Rows(0).Item("Data")
                End If

            End If
        Else
            DPCNAME.Text = DBAdapter1.Rows(0).Item("PCNAME")
        End If

      


        'SetFieldData("PCNAME", DBAdapter1.Rows(0).Item("PCNAME"))

        If Mid(DBAdapter1.Rows(0).Item("ASDate").ToString, 1, 4) = "1900" Then
            DASDate.Text = ""
        Else
            DASDate.Text = DBAdapter1.Rows(0).Item("ASDate")               '�߷|�ɶ�
        End If

        If Mid(DBAdapter1.Rows(0).Item("SignDate").ToString, 1, 4) = "1900" Then
            DSignDate.Text = ""
        Else
            DSignDate.Text = DBAdapter1.Rows(0).Item("SignDate")             'ñ�֮ɶ�
        End If


        '�߷|�ɶ�


        If DBAdapter1.Rows(0).Item("CheckASDate") = "1" Then
            CheckASDate.Items(1).Selected = True
        Else
            CheckASDate.Items(0).Selected = True
        End If



        '���Ӹ��
        SQL = "Select * From F_DISPOSALSheetdt "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

        Dim i As Integer
        Dim ITEMCODE, ITEMNAME, PIECE, METER, YARD, KG, PRICE, DISRULE, REMARK As String

        If DBAdapter2.Rows.Count > 0 Then

            For i = 0 To DBAdapter2.Rows.Count - 1

                ITEMCODE = "DITEMCODE" + Trim(Str(i + 1))
                Dim DITEMCODE As TextBox = Me.FindControl(ITEMCODE)
                If DBAdapter2.Rows(i).Item("ITEMCODE") <> "" Then
                    DITEMCODE.Text = DBAdapter2.Rows(i).Item("ITEMCODE")
                End If


                ITEMNAME = "DITEMNAME" + Trim(Str(i + 1))
                Dim DITEMNAME As TextBox = Me.FindControl(ITEMNAME)
                If DBAdapter2.Rows(i).Item("ITEMNAME") <> "" Then
                    DITEMNAME.Text = DBAdapter2.Rows(i).Item("ITEMNAME")
                End If

                PIECE = "DPIECE" + Trim(Str(i + 1))
                Dim DPIECE As TextBox = Me.FindControl(PIECE)
                If DBAdapter2.Rows(i).Item("PIECE") <> "" Then
                    DPIECE.Text = DBAdapter2.Rows(i).Item("PIECE")
                    DPIECE.Style.Add("TEXT-ALIGN", "right")
                    Dim PRICE1 As Double
                    PRICE1 = DBAdapter2.Rows(i).Item("PIECE")

                    DPIECE.Text = PRICE1.ToString("#,0.00")
                End If


                METER = "DMETER" + Trim(Str(i + 1))
                Dim DMETER As TextBox = Me.FindControl(METER)
                If DBAdapter2.Rows(i).Item("METER") <> "" Then
                    DMETER.Text = DBAdapter2.Rows(i).Item("METER")
                    DMETER.Style.Add("TEXT-ALIGN", "right")
                    Dim METER1 As Double
                    METER1 = DBAdapter2.Rows(i).Item("METER")

                    DMETER.Text = METER1.ToString("#,0.00")
                End If


                YARD = "DYARD" + Trim(Str(i + 1))
                Dim DYARD As TextBox = Me.FindControl(YARD)
                If DBAdapter2.Rows(i).Item("YARD") <> "" Then
                    DYARD.Text = DBAdapter2.Rows(i).Item("YARD")
                    DYARD.Style.Add("TEXT-ALIGN", "right")
                    Dim YARD1 As Double
                    YARD1 = DBAdapter2.Rows(i).Item("YARD")

                    DYARD.Text = YARD1.ToString("#,0.00")
                End If

                KG = "DKG" + Trim(Str(i + 1))
                Dim DKG As TextBox = Me.FindControl(KG)
                If DBAdapter2.Rows(i).Item("KG") <> "" Then
                    DKG.Text = DBAdapter2.Rows(i).Item("KG")
                    DKG.Style.Add("TEXT-ALIGN", "right")
                    Dim KG1 As Double
                    KG1 = DBAdapter2.Rows(i).Item("KG")

                    DKG.Text = KG1.ToString("#,0.00")
                End If


                PRICE = "DPRICE" + Trim(Str(i + 1))
                Dim DPRICE As TextBox = Me.FindControl(PRICE)
                If DBAdapter2.Rows(i).Item("PRICE") <> "" Then
                    DPRICE.Text = DBAdapter2.Rows(i).Item("PRICE")
                    DPRICE.Style.Add("TEXT-ALIGN", "right")
                    Dim DPRICE2 As Double
                    DPRICE2 = DBAdapter2.Rows(i).Item("PRICE")

                    DPRICE.Text = DPRICE2.ToString("#,0.00")
                End If


                DISRULE = "DDISRULE" + Trim(Str(i + 1))
                Dim DDISRULE As DropDownList = Me.FindControl(DISRULE)
                If DBAdapter2.Rows(i).Item("DISRULE") <> "" Then
                    SetFieldData(DISRULE, DBAdapter2.Rows(i).Item("DISRULE"))
                    DDISRULE.SelectedValue = DBAdapter2.Rows(i).Item("DISRULE")
                End If



                REMARK = "DREMARK" + Trim(Str(i + 1))
                Dim DREMARK As TextBox = Me.FindControl(REMARK)
                If DBAdapter2.Rows(i).Item("REMARK") <> "" Then
                    DREMARK.Text = DBAdapter2.Rows(i).Item("REMARK")

                End If


            Next

        End If
        GetTotal() '���s�p��

        Dim SIGN, SIGNDATE As String
        '�g�Jñ�֥D�޸��
        SQL = " select rank() over(order by [step]) as cno,decidename,convert(char(10),completedtime,111)CompletedDate "
        SQL = SQL & " from (select  step,decidename,max(completedtime)completedtime  from  t_waithandle"
        SQL = SQL & " where  FORMNO = '006001' AND  FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & " And Step <= 100  And Step not in (1,500,999)  and sts =1 and flowtype <> 0"
        SQL = SQL & " group by  step,decidename"
        SQL = SQL & " )a order by step "
        Dim DBAdapter4 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter4.Rows.Count > 0 Then

            For i = 0 To DBAdapter4.Rows.Count - 1
                'ñ�֥D��
                SIGN = "DSIGN" + Trim(Str(i + 1))
                Dim DSIGN As TextBox = Me.FindControl(SIGN)
                DSIGN.Text = DBAdapter4.Rows(i).Item("decidename")
                'ñ�֤��
                SIGNDATE = "DSIGNDATE" + Trim(Str(i + 1))
                Dim DSIGNDATE As TextBox = Me.FindControl(SIGNDATE)
                DSIGNDATE.Text = DBAdapter4.Rows(i).Item("CompletedDate")

            Next
        End If



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     ��������ܤ�����J�ˬd
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '�ظm�����ݩʰ}�C
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '��渹�X,�u�{���d���X,���W,����ݩ�

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
        rqdVal.Display = ValidatorDisplay.None              ' �]�b�����W�[�JValidationSummary , �G���ұ���Τ@���
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)



    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer



        idx = FindFieldInf(pFieldName)



        '�Ӷ����o�ǫhappend
        If pFieldName = "RULE" Then
            Dim Rule As String

            For i = 1 To 10
                Rule = "DDISRULE" + Trim(Str(i))
                Dim DText As DropDownList = Me.FindControl(Rule)
                DText.Items.Clear()
            Next

            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    For i = 1 To 10
                        Rule = "DDISRULE" + Trim(Str(i))
                        Dim DText As DropDownList = Me.FindControl(Rule)
                        DText.Items.Add(ListItem1)
                    Next

                End If
            Else
                SQL = "  Select   rank() over(order by [data]) as cno,* from M_referp"
                SQL = SQL & " where  cat = '6001'"
                SQL = SQL & " and dkey = 'RULE'"
                SQL = SQL & " order by unique_id"


                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                '�s�WRULE
                For Each dtr As Data.DataRow In DBAdapter1.Rows

                    For i = 1 To 10
                        Rule = "DDISRULE" + Trim(Str(i))
                        Dim DText As DropDownList = Me.FindControl(Rule)
                        If dtr("cno") = "1" Then
                            DText.Items.Clear()
                            DText.Items.Add("")
                        End If
                        DText.Items.Add(dtr("Data"))
                    Next

                Next

            End If
        End If

        '�Ӷ����o�ǫhshowfomrdata  


        If Mid(pFieldName, 1, 8) = "DDISRULE" Then
            Dim Rule As String
            Dim COUNTS As Integer
            COUNTS = Mid(pFieldName, 9, CInt(pFieldName.Length) - 1)

            i = COUNTS
            Rule = "DDISRULE" + Trim(Str(i))

            Dim DText As DropDownList = Me.FindControl(Rule)
            DText.Items.Clear()

            SQL = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            SQL = SQL & " where  cat = '6001'"
            SQL = SQL & " and dkey = 'RULE'"
            SQL = SQL & " and data ='" + pName + "'"
            SQL = SQL & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            '�s�WRULE
            For Each dtr As Data.DataRow In DBAdapter1.Rows


                Rule = "DDISRULE" + Trim(Str(i))
                DText.Items.Add(dtr("Data"))

            Next


        End If




    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     �M��������ݩ�
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i <= 60 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function

    Sub GetTotal()
        Dim I As Integer
        Dim SUMPIECE1, SUMPIECE2, SUMMETER1, SUMMETER2, SUMYARD1, SUMYARD2, SUMKG1, SUMKG2, SUMPRICE1, SUMPRICE2 As Double
        Dim PIECE, METER, YARD, KG, PRICE, DISRULE As String

        For I = 1 To 10


            PIECE = "DPIECE" + Trim(Str(I))
            Dim DPIECE As TextBox = Me.FindControl(PIECE)

            METER = "DMETER" + Trim(Str(I))
            Dim DMETER As TextBox = Me.FindControl(METER)

            YARD = "DYARD" + Trim(Str(I))
            Dim DYARD As TextBox = Me.FindControl(YARD)

            KG = "DKG" + Trim(Str(I))
            Dim DKG As TextBox = Me.FindControl(KG)

            PRICE = "DPRICE" + Trim(Str(I))
            Dim DPRICE As TextBox = Me.FindControl(PRICE)


            DISRULE = "DDISRULE" + Trim(Str(I))
            Dim DDISRULE As DropDownList = Me.FindControl(DISRULE)


            If DDISRULE.SelectedValue = "�C���k" Then
                If DPIECE.Text <> "" Then
                    SUMPIECE1 = SUMPIECE1 + DPIECE.Text
                End If
                If DMETER.Text <> "" Then
                    SUMMETER1 = SUMMETER1 + DMETER.Text
                End If
                If DYARD.Text <> "" Then
                    SUMYARD1 = SUMYARD1 + DYARD.Text
                End If
                If DKG.Text <> "" Then
                    SUMKG1 = SUMKG1 + DKG.Text
                End If
                If DPRICE.Text <> "" Then
                    SUMPRICE1 = SUMPRICE1 + DPRICE.Text
                End If

            ElseIf DDISRULE.SelectedValue = "�D�C���k" Then
                If DPIECE.Text <> "" Then
                    SUMPIECE2 = SUMPIECE2 + DPIECE.Text
                End If
                If DMETER.Text <> "" Then
                    SUMMETER2 = SUMMETER2 + DMETER.Text
                End If
                If DYARD.Text <> "" Then
                    SUMYARD2 = SUMYARD2 + DYARD.Text
                End If
                If DKG.Text <> "" Then
                    SUMKG2 = SUMKG2 + DKG.Text
                End If
                If DPRICE.Text <> "" Then
                    SUMPRICE2 = SUMPRICE2 + DPRICE.Text
                End If
            End If
        Next

     
        DPIECE11.Text = SUMPIECE1.ToString("#,0.00")
        DPIECE11.Style.Add("TEXT-ALIGN", "right")
        DPIECE12.Text = SUMPIECE2.ToString("#,0.00")
        DPIECE12.Style.Add("TEXT-ALIGN", "right")
        DPIECE12.Text = SUMPIECE2.ToString("#,0.00")
        DMETER11.Text = SUMMETER1.ToString("#,0.00")
        DMETER11.Style.Add("TEXT-ALIGN", "right")
        DMETER12.Text = SUMMETER2.ToString("#,0.00")
        DMETER12.Style.Add("TEXT-ALIGN", "right")
        DYARD11.Text = SUMYARD1.ToString("#,0.00")
        DYARD11.Style.Add("TEXT-ALIGN", "right")
        DYARD12.Text = SUMYARD2.ToString("#,0.00")
        DYARD12.Style.Add("TEXT-ALIGN", "right")
        DKG11.Text = SUMKG1.ToString("#,0.00")
        DKG11.Style.Add("TEXT-ALIGN", "right")
        DKG12.Text = SUMKG2.ToString("#,0.00")
        DKG12.Style.Add("TEXT-ALIGN", "right")
        DKG11.Text = SUMKG1.ToString("#,0.00")

        DPRICE11.Text = SUMPRICE1.ToString("#,0.00")
        DPRICE11.Style.Add("TEXT-ALIGN", "right")
        DPRICE12.Text = SUMPRICE2.ToString("#,0.00")
        DPRICE12.Style.Add("TEXT-ALIGN", "right")
        DPIECE13.Text = (SUMPIECE1 + SUMPIECE2).ToString("#,0.00")
        DPRICE13.Style.Add("TEXT-ALIGN", "right")
        DMETER13.Text = (SUMMETER1 + SUMMETER2).ToString("#,0.00")
        DMETER13.Style.Add("TEXT-ALIGN", "right")
        DYARD13.Text = (SUMYARD1 + SUMYARD2).ToString("#,0.00")
        DYARD13.Style.Add("TEXT-ALIGN", "right")
        DKG13.Text = (SUMKG1 + SUMKG2).ToString("#,0.00")
        DKG13.Style.Add("TEXT-ALIGN", "right")
        DPRICE13.Text = (SUMPRICE1 + SUMPRICE2).ToString("#,0.00")
        DPIECE13.Style.Add("TEXT-ALIGN", "right")

    End Sub

  
End Class

