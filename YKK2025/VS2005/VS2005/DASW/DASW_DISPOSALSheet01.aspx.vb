Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls



 


Partial Class DASW_DISPOSALSheet01
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
    Dim griderror, inserterror, fielderror As Boolean

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�
        TopPosition()           '���s��RequestedField��Top��m
        SetControlPosition()    ' �]�w�����m

        If Not Me.IsPostBack Then   '���OPostBack
            ShowSheetField("New")   '��������ܤ�����J�ˬd
            ShowSheetFunction()     '���\����s���

            If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
                ShowFormData()      '��ܪ����
                UpdateTranFile()    '��s������
                TopPosition()
                SetControlPosition()    ' �]�w�����m

            End If
            SetPopupFunction()      '�]�w�u�X�����ƥ�


        Else
            ShowSheetFunction()     '���\����s���
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' �]�w�w�]��ñ�֪�
            ShowSheetField("Post")  '��������ܤ�����J�ˬd

            ShowMessage()           '�W�Ǹ���ˬd����ܰT��

            '�W�Ǹ���ˬd����ܰT��

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
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�

        wbFormSno = Request.QueryString("pFormSno")  '�s��_��άy����
        wbStep = Request.QueryString("pStep")        '�s��_��Τu�{�N�X
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID

        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        'Add-Start by Joy 2009/11/20(2010��ƾ����)
        '�ӽЪ�-�s�զ�ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)

        ' wUserID = Request.QueryString("pUserID")


        ' wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))

        Response.Cookies("PGM").Value = "DASW_DISPOSALSheet01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X

       
        ' BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BASDate.Attributes.Add("onClick", "calendarPicker('DASDate')")
        BDISPOSALTYPE.Attributes("onclick") = "GetDISPOSALTYPE();"

        BReason.Attributes.Add("onclick", "window.open('DASW_DISPOSALReasonURL.aspx','','status=0,toolbar=0,width=1000,height=600')")


    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��s������
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�
        'Modify-End
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
        SQL = "Select *,convert(char(10),SignDate,111)SignDate1,convert(char(10),AppDate,111)Appdate1   From F_DISPOSALSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows(0).Item("MKTSIGN") = 1 Then
            DMKTSIGN.Checked = True
        Else
            DMKTSIGN.Checked = False
        End If

        If DBAdapter1.Rows(0).Item("CUSTOMERTOLL") = 1 Then
            DCUSTOMERTOLL.Checked = True
        Else
            DCUSTOMERTOLL.Checked = False
        End If

        DSignDate.Text = DBAdapter1.Rows(0).Item("SignDate1")
        DDisposalYM.Text = DBAdapter1.Rows(0).Item("DisposalYM")
        DNo.Text = DBAdapter1.Rows(0).Item("No")
        DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate1")
        DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")
        DAppName.Text = DBAdapter1.Rows(0).Item("AppName")
        SetFieldData("DISPOSALREASON", DBAdapter1.Rows(0).Item("DISPOSALREASON"))
        DDISPOSALREASON1.Text = DBAdapter1.Rows(0).Item("DISPOSALREASON1")
        SetFieldData("DUTYDEPO", DBAdapter1.Rows(0).Item("DUTYDEPO"))
        DSales.Text = DBAdapter1.Rows(0).Item("Sales")
        SetFieldData("PLACE", DBAdapter1.Rows(0).Item("PLACE"))
        SetFieldData("STYPE", DBAdapter1.Rows(0).Item("STYPE"))
        SetFieldData("DISPOSALRULE", DBAdapter1.Rows(0).Item("DISPOSALRULE"))
        DDISPOSALTYPE.Value = DBAdapter1.Rows(0).Item("DISPOSALTYPE")
        DCHINESEREASON.Text = DBAdapter1.Rows(0).Item("CHINESEREASON")
        DJAPANREASON.Text = DBAdapter1.Rows(0).Item("JAPANREASON")
        If DBAdapter1.Rows(0).Item("DISPOSALFILE1") <> "" Then
            LDISPOSALFILE.NavigateUrl = Path & DBAdapter1.Rows(0).Item("DISPOSALFILE1") '���o����
            LDISPOSALFILE.Visible = True
        Else
            LDISPOSALFILE.Visible = False
        End If

        DPCNAME.Text = DBAdapter1.Rows(0).Item("PCNAME")

        'SetFieldData("PCNAME", DBAdapter1.Rows(0).Item("PCNAME"))


        If Mid(DBAdapter1.Rows(0).Item("SignDate").ToString, 1, 4) = "1900" Then
            DSignDate.Text = ""
        Else
            DSignDate.Text = DBAdapter1.Rows(0).Item("SignDate1")             'ñ�֮ɶ�
        End If

        '�߷|�ɶ�

        Dim j As Integer
        For j = 0 To CheckASDate.Items.Count - 1 Step j + 1
            ' check item selected state.
            If DBAdapter1.Rows(0).Item("CheckASDate") = CheckASDate.Items(j).Value Then
                CheckASDate.Items(j).Selected = True
            End If
        Next

        If wStep = 110 Or wStep = 30 Then
            If DBAdapter1.Rows(0).Item("CheckASDate") Then
                DASDate.Style.Add("background-color", "greenyellow")
            Else
                DASDate.Style.Add("background-color", " yellow")
            End If


        End If

     
        If wStep = 500 Then  '�ͺޥD��KEY�߷|���

            If DSignDate.Text <> "�L" Then

                If DSignDate.Text <> "���}��" Then
                    Dim SQL1 As String
                    '�D�ɸ��
                    SQL1 = " SELECT  *  FROM ( "
                    SQL1 = SQL1 + " SELECT  SUBSTRING(DKEY,13,1) AS  DKEYS, DATA AS SDATE   from M_referp where(cat = 6001) and left(dkey,12)='SendTime-PCS'   )A, ("
                    SQL1 = SQL1 + "  SELECT SUBSTRING(DKEY,12,1) AS  DKEYE,DATA AS ACCDATE   from M_referp  where(cat = 6001) and left(dkey,11)='AccountTime'   )B  "
                    SQL1 = SQL1 + " WHERE (A.DKEYS = B.DKEYE)  and sdate = convert(char(10), dateadd(day,+1,convert(datetime,'" + DSignDate.Text + "')),111)"

                    Dim DBAdapter5 As DataTable = uDataBase.GetDataTable(SQL1)
                    If DBAdapter5.Rows.Count > 0 Then
                        DASDate.Value = DBAdapter5.Rows(0).Item("ACCDATE").ToString
                        If DASDate.Value <> "" Then
                            CheckASDate.Items(0).Selected = False
                            CheckASDate.Items(1).Selected = True
                        Else
                            CheckASDate.Items(0).Selected = True
                            CheckASDate.Items(1).Selected = False
                        End If


                    End If
                End If

            End If

          

        Else
            If Mid(DBAdapter1.Rows(0).Item("ASDate").ToString, 1, 4) = "1900" Then
                DASDate.Value = ""
            Else
                DASDate.Value = DBAdapter1.Rows(0).Item("ASDate")               '�߷|�ɶ�
            End If
        End If




        'Dim SQL As String
        '�D�ɸ��
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '6001'"
        SQL = SQL + " and dkey ='DisposalFilePath'"
        Dim DBAdapter6 As DataTable = uDataBase.GetDataTable(SQL)
        Dim OpenDir1 As String = ""
        If DBAdapter6.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter6.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter6.Rows(0).Item("Data1")
        End If


        '�ˬd�ؿ��O�_�s�A���s�b�N�s�ؤ@��
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If


        '�}�ҳ��o��Ƨ����|
        DDISPOSALFILE2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "\" + DNo.Text + "','_blank');return true;")




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

                    DPRICE.Text = DPRICE2.ToString("#,0")
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


        getDUTYDEPOID()





        '�g�J�y�{���
        SQL = "Select * From T_WaitHandle "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
        SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"

        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter3.Rows.Count > 0 Then
            DDecideDesc.Text = DBAdapter3.Rows(0).Item("DecideDesc")       '����


            If DBAdapter3.Rows(0).Item("BEndTime") < NowDateTime Then
                If DDelay.Visible = True Then
                    SetFieldData("ReasonCode", DBAdapter3.Rows(0).Item("ReasonCode"))    '����z�ѥN�X
                    If DBAdapter3.Rows(0).Item("ReasonCode") = "" Then
                        SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
                        DReasonDesc.Text = ""   '�����L����
                    Else
                        DReason.Text = DBAdapter3.Rows(0).Item("Reason")  '����z��
                        DReasonDesc.Text = DBAdapter3.Rows(0).Item("ReasonDesc")     '�����L����
                    End If
                End If
            End If

        End If






        '�֩w�i�����
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "AEndTimeDesc As Description, "
        SQL = SQL + "URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
        'SQL = SQL + "Order by Unique_ID Desc "
        SQL = SQL + " Order by CreateTime Desc, Step Desc, SeqNo Desc "
        Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet9, "DecideHistory")
        GridView2.DataSource = dtWaitHandle1
        GridView2.DataBind()

        'DB�s������
        'OleDbConnection1.Close()

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     ���\����s���
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String

        'DB�s���]�w
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)


        If DBAdapter1.Rows.Count > 0 Then
            '�q�lñ�����ϥ�
            If DBAdapter1.Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '���[�ɮץ��ϥ�(��FormField���]�w)
            If DBAdapter1.Rows(0).Item("Attach") = 1 Then
            Else
            End If
            '�x�s���s
            If DBAdapter1.Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBAdapter1.Rows(0).Item("SaveDesc")
                BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1���s
            If DBAdapter1.Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBAdapter1.Rows(0).Item("NGDesc1")
                BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                wNGSts1 = DBAdapter1.Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2���s
            If DBAdapter1.Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBAdapter1.Rows(0).Item("NGDesc2")
                BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                wNGSts2 = DBAdapter1.Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK���s
            If DBAdapter1.Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBAdapter1.Rows(0).Item("OKDesc")
                BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                wOKSts = DBAdapter1.Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '��Ǻ޲z
            If DBAdapter1.Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then

                '��Ǻ޲z
                If wDelay = 1 Then
                    If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '����Sheet
                    Else
                        DDelay.Visible = False  '����Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '����z�ѥN�X
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "���`�G�ݿ�J����z��")
                    DReason.Visible = True         '����z��
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "���`�G�ݿ�J����z��")
                    DReasonDesc.Visible = True     '�����L����
                    Top = 1220
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 1220
                End If



                '������
                DDecideDesc.Visible = True      '����
                '�����ݿ�J
                If DBAdapter2.Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "���`�G�ݿ�J����")
                End If

                '���s��m
                BNG1.Style.Add("Top", Top)      'NG���s
                BNG2.Style.Add("Top", Top)     'NG���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s

            End If
        Else
            Top = 1150
            'Sheet����

            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False     '����
            DDescSheet.Visible = False
            DReasonCode.Visible = False     '����z�ѥN�X
            DReason.Visible = False         '����z��
            DReasonDesc.Visible = False     '�����L����
            DHistoryLabel.Visible = False  '�֩w�i��
            '���s��m
            BNG1.Style.Add("Top", Top)      'NG���s
            BNG2.Style.Add("Top", Top)     'NG���s
            BSAVE.Style.Add("Top", Top)    '�x�s���s
            BOK.Style.Add("Top", Top)      'OK���s

        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     �]�w�u�X�����ƥ�
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        '������



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����z�ѥN�X�I���ƥ�
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String


        '���s��RequestedField��Top��m
        If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter1.Rows.Count > 0 Then
                If DBAdapter1.Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 1300
                Else
                    If DDelay.Visible = True Then
                        Top = 1300
                    Else
                        Top = 1300
                    End If
                End If
            End If
        Else
            Top = 1300

        End If
        '----


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

        SetFieldAttribute(pPost)     '���U����ݩʤ�����J�ˬd���]�w
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     ���U����ݩʤ�����J�ˬd���]�w
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '���U����ݩʤ�����J�ˬd���]�w

      
        'No
        Select Case FindFieldInf("NO")
            Case 0  '���
                DNo.BackColor = Color.LightGray
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '�ק�+�ˬd
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "���`�G�ݿ�J�ܢ�")
                DNo.Visible = True
            Case 2  '�ק�
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '����
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""



        '���
        Select Case FindFieldInf("AppDate")
            Case 0  '���
                DAppDate.BackColor = Color.LightGray
                DAppDate.ReadOnly = True
                DAppDate.Visible = True

            Case 1  '�ק�+�ˬd
                DAppDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "���`�G�ݿ�J���")
                DAppDate.Visible = True

            Case 2  '�ק�
                DAppDate.BackColor = Color.Yellow
                DAppDate.Visible = True

            Case Else   '����
                DAppDate.Visible = False

        End Select
        If pPost = "New" Then DAppDate.Text = Now.ToString("yyyy/MM/dd") '�{�b���

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '���
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '�ק�+�ˬd
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "���`�G�ݿ�J����")
                DDepoName.Visible = True
            Case 2  '�ק�
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '����
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DepoName", "ZZZZZZ")


        'Name
        Select Case FindFieldInf("AppName")
            Case 0  '���
                DAppName.BackColor = Color.LightGray
                DAppName.ReadOnly = True
                DAppName.Visible = True
            Case 1  '�ק�+�ˬd
                DAppName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAPPNameRqd", "DAPPName", "���`�G�ݿ�J�m�W")
                DAppName.Visible = True
            Case 2  '�ק�
                DAppName.BackColor = Color.Yellow
                DAppName.Visible = True
            Case Else   '����
                DAppName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AppName", "ZZZZZZ")



        '�ӽЭ�]
        Select Case FindFieldInf("DISPOSALREASON")
            Case 0  '���
                DDISPOSALREASON.BackColor = Color.LightGray
                DDISPOSALREASON.Visible = True

            Case 1  '�ק�+�ˬd
                DDISPOSALREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDISPOSALREASONRqd", "DDISPOSALREASON", "���`�G�ݿ�ӽЭ�]")
                DDISPOSALREASON.Visible = True
            Case 2  '�ק�
                DDISPOSALREASON.BackColor = Color.Yellow
                DDISPOSALREASON.Visible = True
            Case Else   '����
                DDISPOSALREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DISPOSALREASON", "ZZZZZZ")


        '�䥦��]
        Select Case FindFieldInf("DISPOSALREASON1")
            Case 0  '���
                DDISPOSALREASON1.BackColor = Color.LightGray
                DDISPOSALREASON1.ReadOnly = True
                DDISPOSALREASON1.Visible = True
            Case 1  '�ק�+�ˬd
                DDISPOSALREASON1.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DDISPOSALREASON1Rqd", "DDISPOSALREASON1", "���`�G�ݿ�J�䥦��]")
                DDISPOSALREASON1.Visible = True
            Case 2  '�ק�
                DDISPOSALREASON1.BackColor = Color.Yellow
                DDISPOSALREASON1.Visible = True
            Case Else   '����
                DDISPOSALREASON1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DISPOSALREASON1", "ZZZZZZ")




        '�~��
        Select Case FindFieldInf("Sales")
            Case 0  '���
                DSales.BackColor = Color.LightGray
                DSales.ReadOnly = True
                DSales.Visible = True
            Case 1  '�ק�+�ˬd
                DSales.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DSalesRqd", "DSales", "���`�G�ݿ�J�䥦��]")
                DSales.Visible = True
            Case 2  '�ק�
                DSales.BackColor = Color.Yellow
                DSales.Visible = True
            Case Else   '����
                DSales.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sales", "ZZZZZZ")


        '���o�Φ�
        Select Case FindFieldInf("STYPE")
            Case 0  '���
                DSTYPE.BackColor = Color.LightGray
                DSTYPE.Visible = True

            Case 1  '�ק�+�ˬd
                DSTYPE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSTYPERqd", "DSTYPE", "���`�G�ݿ�J���o�Φ�")
                DSTYPE.Visible = True
            Case 2  '�ק�
                DSTYPE.BackColor = Color.Yellow
                DSTYPE.Visible = True
            Case Else   '����
                DSTYPE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("STYPE", "ZZZZZZ")



        '�d������
        Select Case FindFieldInf("DUTYDEPO")
            Case 0  '���
                DDUTYDEPO.BackColor = Color.LightGray
                DDUTYDEPO.Visible = True

            Case 1  '�ק�+�ˬd
                DDUTYDEPO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDUTYDEPORqd", "DDUTYDEPO", "���`�G�ݿ�J�d������")
                DDUTYDEPO.Visible = True
            Case 2  '�ק�
                DDUTYDEPO.BackColor = Color.Yellow
                DDUTYDEPO.Visible = True
            Case Else   '����
                DDUTYDEPO.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DUTYDEPO", "ZZZZZZ")

        '����
        Select Case FindFieldInf("PLACE")
            Case 0  '���
                DPLACE.BackColor = Color.LightGray
                DPLACE.Visible = True

            Case 1  '�ק�+�ˬd
                DPLACE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPLACERqd", "DPLACE", "���`�G�ݿ�J�겣�b�w����")
                DPLACE.Visible = True
            Case 2  '�ק�
                DPLACE.BackColor = Color.Yellow
                DPLACE.Visible = True
            Case Else   '����
                DPLACE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PLACE", "ZZZZZZ")

        '���o�ǫh
        Select Case FindFieldInf("DISPOSALRULE")
            Case 0  '���
                DDISPOSALRULE.BackColor = Color.LightGray
                DDISPOSALRULE.Visible = True

            Case 1  '�ק�+�ˬd
                DDISPOSALRULE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDISPOSALRULERqd", "DDISPOSALRULE", "���`�G�ݿ�J���o�ǫh")
                DDISPOSALRULE.Visible = True
            Case 2  '�ק�
                DDISPOSALRULE.BackColor = Color.Yellow
                DDISPOSALRULE.Visible = True
            Case Else   '����
                DDISPOSALRULE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DISPOSALRULE", "ZZZZZZ")



        'Name
        Select Case FindFieldInf("AppName")
            Case 0  '���
                DAppName.BackColor = Color.LightGray
                DAppName.ReadOnly = True
                DAppName.Visible = True
            Case 1  '�ק�+�ˬd
                DAppName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAPPNameRqd", "DAPPName", "���`�G�ݿ�J�m�W")
                DAppName.Visible = True
            Case 2  '�ק�
                DAppName.BackColor = Color.Yellow
                DAppName.Visible = True
            Case Else   '����
                DAppName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AppName", "ZZZZZZ")




        'DisposalType
        Select Case FindFieldInf("DISPOSALTYPE")
            Case 0  '���
                DDISPOSALTYPE.Visible = True
                DDISPOSALTYPE.Style.Add("background-color", "lightgrey")
                DDISPOSALTYPE.Attributes.Add("readonly", "true")
                BDISPOSALTYPE.Visible = False

            Case 1  '�ק�+�ˬd
                DDISPOSALTYPE.Visible = True
                DDISPOSALTYPE.Style.Add("background-color", "greenyellow")
                DDISPOSALTYPE.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DDISPOSALTYPERqd", "DDISPOSALTYPE", "���`�G�ݿ�J���o�~���O")
            Case 2  '�ק�
                DDISPOSALTYPE.Visible = True
                DDISPOSALTYPE.Style.Add("background-color", "yellow")
                DDISPOSALTYPE.Attributes.Add("readonly", "true")
                BDISPOSALTYPE.Visible = True
            Case Else   '����
                DDISPOSALTYPE.Visible = False
                BDISPOSALTYPE.Visible = False

        End Select
        If pPost = "New" Then DDISPOSALTYPE.Value = ""





        'CHINESEREASON
        Select Case FindFieldInf("CHINESEREASON")
            Case 0  '���
                DCHINESEREASON.BackColor = Color.LightGray
                DCHINESEREASON.ReadOnly = True
                DCHINESEREASON.Visible = True
            Case 1  '�ק�+�ˬd
                DCHINESEREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCHINESEREASONRqd", "DCHINESEREASON", "���`�G�ݿ�J�����]")
                DCHINESEREASON.Visible = True
            Case 2  '�ק�
                DCHINESEREASON.BackColor = Color.Yellow
                DCHINESEREASON.Visible = True
            Case Else   '����
                DCHINESEREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CHINESEREASON", "ZZZZZZ")


        'JAPANREASON
        Select Case FindFieldInf("JAPANREASON")
            Case 0  '���
                DJAPANREASON.BackColor = Color.LightGray
                DJAPANREASON.ReadOnly = True
                DJAPANREASON.Visible = True
            Case 1  '�ק�+�ˬd
                DJAPANREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJAPANREASONRqd", "DJAPANREASON", "���`�G�ݿ�J����]")
                DJAPANREASON.Visible = True
            Case 2  '�ק�
                DJAPANREASON.BackColor = Color.Yellow
                DJAPANREASON.Visible = True
            Case Else   '����
                DJAPANREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("JAPANREASON", "ZZZZZZ")


        Dim i As Integer
        Dim ITEMCODE As String
        Select Case FindFieldInf("ITEMCODE")
            Case 0  '���
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DITEMCODE" + Str(i) + "Rqd", ITEMCODE, "���`�G�ݿ�JITEMCODE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 10
            ITEMCODE = "DITEMCODE" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(ITEMCODE)
            If pPost = "New" Then SetFieldData(ITEMCODE, "ZZZZZZ")
        Next

        Dim ITEMNAME As String
        Select Case FindFieldInf("ITEMNAME")
            Case 0  '���
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DITEMNAME" + Str(i) + "Rqd", ITEMNAME, "���`�G�ݿ�JITEMNAME" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.Visible = False
                Next

        End Select


        For i = 1 To 13
            ITEMNAME = "DITEMNAME" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(ITEMNAME)
            If pPost = "New" Then SetFieldData(ITEMNAME, "ZZZZZZ")
        Next

        Dim PIECE As String
        Select Case FindFieldInf("PIECE")
            Case 0  '���
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DPIECE" + Str(i) + "Rqd", PIECE, "���`�G�ݿ�JPIECE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 13
            PIECE = "DPIECE" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(PIECE)
            If pPost = "New" Then SetFieldData(PIECE, "ZZZZZZ")
            DText.Attributes("onkeyup") = "trans_amt('" + DText.ClientID + "');"
        Next

        Dim METER As String
        Select Case FindFieldInf("METER")
            Case 0  '���
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DMETER" + Str(i) + "Rqd", METER, "���`�G�ݿ�JMETER" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 13
            METER = "DMETER" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(METER)
            If pPost = "New" Then SetFieldData(METER, "ZZZZZZ")
            DText.Attributes("onkeyup") = "trans_amt('" + DText.ClientID + "');"
        Next


        Dim YARD As String
        Select Case FindFieldInf("YARD")
            Case 0  '���
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DYARD" + Str(i) + "Rqd", YARD, "���`�G�ݿ�JYARD" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 13
            YARD = "DYARD" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(YARD)
            If pPost = "New" Then SetFieldData(YARD, "ZZZZZZ")
            DText.Attributes("onkeyup") = "trans_amt('" + DText.ClientID + "');"
        Next


        Dim KG As String
        Select Case FindFieldInf("KG")
            Case 0  '���
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DKG" + Str(i) + "Rqd", KG, "���`�G�ݿ�JKG" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 13
            KG = "DKG" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(KG)
            If pPost = "New" Then SetFieldData(KG, "ZZZZZZ")
            DText.Attributes("onkeyup") = "trans_amt('" + DText.ClientID + "');"
        Next





        Dim PRICE As String
        Select Case FindFieldInf("PRICE")
            Case 0  '���
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DPRICE" + Str(i) + "Rqd", PRICE, "���`�G�ݿ�JPRICE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 13
            PRICE = "DPRICE" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(PRICE)
            If pPost = "New" Then SetFieldData(PRICE, "ZZZZZZ")
            DText.Attributes("onkeyup") = "trans_amt('" + DText.ClientID + "');"
        Next


        'SUM
        Select Case FindFieldInf("SUM")
            Case 0  '���
                DSUM.Visible = True
            Case 1  '�ק�+�ˬd

                DSUM.Visible = True
            Case 2  '�ק�

                DSUM.Visible = True
            Case Else   '����
                DSUM.Visible = False
        End Select



        Dim DDISRULE As String
        Select Case FindFieldInf("RULE")
            Case 0  '���
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.BackColor = Color.LightGray
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDISRULE " + Str(i) + "Rqd", DDISRULE, "���`�G�ݿ�J���o�W�h" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 10
            DDISRULE = " DDISRULE " + Trim(Str(i))
            Dim DText As DropDownList = Me.FindControl(DDISRULE)
            If pPost = "New" Then SetFieldData("RULE", "ZZZZZZ")
        Next



        Dim REMARK As String
        Select Case FindFieldInf("REMARK")
            Case 0  '���
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DREMARK" + Str(i) + "Rqd", REMARK, "���`�G�ݿ�JREMARK" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 10
            REMARK = "DREMARK" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(REMARK)
            If pPost = "New" Then SetFieldData(REMARK, "ZZZZZZ")
        Next




        '�ͺަ��b
        Select Case FindFieldInf("PCNAME")
            Case 0  '���
                DPCNAME.BackColor = Color.LightGray
                DPCNAME.Visible = True
                DPCNAME.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DPCNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPCNAMERqd", "DPCNAME", "���`�G�ݿ�J�ͺަ��b���H��")
                DPCNAME.Visible = True
            Case 2  '�ק�
                DPCNAME.BackColor = Color.Yellow
                DPCNAME.Visible = True
            Case Else   '����
                DPCNAME.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PCNAME", "ZZZZZZ")


        '���o����
        Select Case FindFieldInf("DISPOSALFILE1")
            Case 0  '���
                DFileUpload1.Visible = False
                DFileUpload1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ' ShowRequiredFieldValidator("DFileUpload1Rqd", "DFileUpload1", "���`�G�ݤW�ǳ��o����")
                DFileUpload1.Visible = True
                DFileUpload1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DFileUpload1.Visible = True
                DFileUpload1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DFileUpload1.Visible = False
        End Select


        '���o����
        Select Case FindFieldInf("DISPOSALFILE2")
            Case 0  '���
                DDISPOSALFILE2.Enabled = False
            Case 1  '�ק�+�ˬd

                DDISPOSALFILE2.Visible = True

            Case 2  '�ק�
                DDISPOSALFILE2.Visible = True

            Case Else   '����
                DDISPOSALFILE2.Visible = False
        End Select



        'ñ�֥D��
        Dim SIGN As String
        Select Case FindFieldInf("SIGN")
            Case 0  '���
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DSIGN" + Str(i) + "Rqd", SIGN, "���`�G�ݿ�JSIGN" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 12
            SIGN = "DSIGN" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(SIGN)
            If pPost = "New" Then SetFieldData(SIGN, "ZZZZZZ")
        Next


        'ñ�֥D��
        Dim SIGNDATE As String
        Select Case FindFieldInf("SIGNDATE")
            Case 0  '���
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '�ק�+�ˬd
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DSIGNDATE" + Str(i) + "Rqd", SIGNDATE, "���`�G�ݿ�JSIGNDATE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '�ק�
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '����
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.Visible = False
                Next

        End Select

        For i = 1 To 12
            SIGNDATE = "DSIGNDATE" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(SIGNDATE)
            If pPost = "New" Then SetFieldData(SIGNDATE, "ZZZZZZ")
        Next




        '��~ñ��
        Select Case FindFieldInf("MKTSIGN")
            Case 0  '���
                ' DMKTSIGN.BackColor = Color.LightGray
                DMKTSIGN.Enabled = False

            Case 1  '�ק�+�ˬd
                DMKTSIGN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMKTSIGNRqd", "DMKTSIGN", "���`�G�ݿ�J����~ñ��")

            Case 2  '�ק�
                DMKTSIGN.BackColor = Color.LightGreen

            Case Else   '����
                DMKTSIGN.Visible = False
        End Select
        If pPost = "New" Then DMKTSIGN.Checked = False


        '�V�Ȥ����
        Select Case FindFieldInf("CUSTOMERTOLL")
            Case 0  '���
                ' DCUSTOMERTOLL.BackColor = Color.LightGray
                DCUSTOMERTOLL.Enabled = False

            Case 1  '�ק�+�ˬd
                DCUSTOMERTOLL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCUSTOMERTOLLRqd", "DCUSTOMERTOLL", "���`�G�ݿ�J�V�Ȥ����")

            Case 2  '�ק�
                DCUSTOMERTOLL.BackColor = Color.LightGreen

            Case Else   '����
                DCUSTOMERTOLL.Visible = False
        End Select
        If pPost = "New" Then DCUSTOMERTOLL.Checked = False



        '�߷|���
        'DeliveryDate
        Select Case FindFieldInf("ASDate")
            Case 0  '���
                DASDate.Visible = True
                DASDate.Style.Add("background-color", "lightgrey")
                DASDate.Attributes.Add("readonly", "true")
                BASDate.Visible = False

            Case 1  '�ק�+�ˬd
                DASDate.Visible = True
                DASDate.Style.Add("background-color", "greenyellow")
                DASDate.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DASDateRqd", "DASDate", "���`�G�ݿ�J�߷|���")
                BASDate.Visible = True

            Case 2  '�ק�
                DASDate.Visible = True
                DASDate.Style.Add("background-color", "yellow")
                DASDate.Attributes.Add("readonly", "true")
                BASDate.Visible = True
            Case Else   '����
                DASDate.Visible = False
                BASDate.Visible = False

        End Select
        If pPost = "New" Then DASDate.Value = ""



        'If pPost = "New" Then DASDate.Value = Now.ToString("yyyy/MM/dd") '�{�b���

        '�߷|���
        Select Case FindFieldInf("SignDate")
            Case 0  '���
                DSignDate.BackColor = Color.LightGray
                DSignDate.ReadOnly = True
                DSignDate.Visible = True

            Case 1  '�ק�+�ˬd
                DSignDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSignDateRqd", "DSignDate", "���`�G�ݿ�Jñ�֤��")
                DSignDate.Visible = True

            Case 2  '�ק�
                DSignDate.BackColor = Color.Yellow
                DSignDate.Visible = True

            Case Else   '����
                DSignDate.Visible = False

        End Select
        'If pPost = "New" Then DSignDate.Text = Now.ToString("yyyy/MM/dd") '�{�b���


        '�O�_�ݥ߷|
        Select Case FindFieldInf("CheckASDate")
            Case 0  '���
                ' DCUSTOMERTOLL.BackColor = Color.LightGray
                CheckASDate.Enabled = False

            Case 1  '�ק�+�ˬd
                CheckASDate.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("CblRqd", "Cbl", "���`�G�ݿ�J�O�_�ݥ߷|")

            Case 2  '�ק�
                CheckASDate.BackColor = Color.Yellow

            Case Else   '����
                CheckASDate.Visible = False
        End Select
        'If pPost = "New" Then cbl_1.Checked = False

        If wStep = 1 Then  '�ͺޥD��KEY�߷|���
            If DSignDate.Text <> "�L" Then
                If DSignDate.Text <> "���}��" Then
                    Dim SQL1 As String
                    '�D�ɸ��
                    SQL1 = " SELECT  *  FROM ( "
                    SQL1 = SQL1 + " SELECT  SUBSTRING(DKEY,13,1) AS  DKEYS, DATA AS SDATE   from M_referp where(cat = 6001) and left(dkey,12)='SendTime-PCS'   )A, ("
                    SQL1 = SQL1 + "  SELECT SUBSTRING(DKEY,12,1) AS  DKEYE,DATA AS ACCDATE   from M_referp  where(cat = 6001) and left(dkey,11)='AccountTime'   )B  "
                    SQL1 = SQL1 + " WHERE (A.DKEYS = B.DKEYE)  and sdate < convert(char(10), dateadd(day,+1,convert(datetime,'" + DSignDate.Text + "')),111)"

                    Dim DBAdapter5 As DataTable = uDataBase.GetDataTable(SQL1)
                    If DBAdapter5.Rows.Count > 0 Then
                        DASDate.Value = DBAdapter5.Rows(0).Item("ACCDATE").ToString
                        If DASDate.Value <> "" Then
                            CheckASDate.Items(0).Selected = False
                            CheckASDate.Items(1).Selected = True
                        Else
                            CheckASDate.Items(0).Selected = True
                            CheckASDate.Items(1).Selected = False
                        End If

                    End If
                End If
            End If



        End If


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

        '���̤γ��� 

        ' SQL = "Select Divname,Username From M_Users "
        ' SQL = SQL & " Where UserID = '" & wApplyID & "'"
        ' SQL = SQL & "   And Active = '1' "



        SQL = "select empname AS  Username,depname AS Divname from ("
        SQL = SQL & "  select empname,depname from M_WFSEMP"
        SQL = SQL & " UNION ALL "
        SQL = SQL & " select empname,depname  from M_WFSEMPcl"
        SQL = SQL & " )a ,M_Users b where a.empname =b.username "
        SQL = SQL & " and UserID = '" & wApplyID & "'"
        SQL = SQL & "  And Active = '1' "

        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)

        DDepoName.Text = DBUser.Rows(0).Item("Divname")
        DAppName.Text = DBUser.Rows(0).Item("Username")


        '���o���b�H��
        DPCNAME.Text = GetPCMAN(DAppName.Text)


        'ñ�ִ���
        If DSTYPE.SelectedValue = "�릸" Or wStep = 1 Then

            If checkDate() = 1 Then
                DSignDate.Text = SignDate

            Else
                If wStep = 1 Then
                    DSignDate.Text = "���}��"
                End If

            End If

            If DSTYPE.SelectedValue = "�g������" Then
                DSignDate.Text = "�L"
            End If

        End If




        '�ӽЭ�]
        If pFieldName = "DISPOSALREASON" Then
            DDISPOSALREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDISPOSALREASON.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '6001'"
                SQL = SQL & " and dkey = 'DISPOSALREASON'"
                SQL = SQL & " order by data"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDISPOSALREASON.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDISPOSALREASON.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '�d������
        If pFieldName = "DUTYDEPO" Then
            DDUTYDEPO.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDUTYDEPO.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '6001'"
                SQL = SQL & " and dkey = 'DUTYDEPO'"
                SQL = SQL & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDUTYDEPO.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDUTYDEPO.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '�b�w����
        If pFieldName = "PLACE" Then
            DPLACE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPLACE.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '6001'"
                SQL = SQL & " and dkey = 'PLACE'"
                SQL = SQL & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DPLACE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPLACE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '���o����
        If pFieldName = "STYPE" Then
            DSTYPE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSTYPE.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '6001'"
                SQL = SQL & " and dkey = 'STYPE'"
                SQL = SQL & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DSTYPE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSTYPE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If



        '���o�ǫh
        If pFieldName = "DISPOSALRULE" Then
            DDISPOSALRULE.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDISPOSALRULE.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '6001'"
                SQL = SQL & " and dkey = 'DISPOSALRULE'"
                SQL = SQL & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDISPOSALRULE.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDISPOSALRULE.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


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
            ' SQL = SQL & " and data ='" + pName + "'"
            SQL = SQL & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            '�s�WRULE
            For Each dtr As Data.DataRow In DBAdapter1.Rows


                Rule = "DDISRULE" + Trim(Str(i))
                DText.Items.Add(dtr("Data"))

            Next


        End If



        ''�ͺަ��b�H��
        'If pFieldName = "PCNAME" Then
        '    DPCNAME.Items.Clear()
        '    If idx = 0 Then
        '        If pName <> "ZZZZZZ" Then
        '            Dim ListItem1 As New ListItem
        '            ListItem1.Text = pName
        '            ListItem1.Value = pName
        '            DPCNAME.Items.Add(ListItem1)
        '        End If
        '    Else
        '        SQL = " Select  a.*,userid  from M_referp a,M_users b"
        '        SQL = SQL & "   where  cat = '6001'"
        '        SQL = SQL & "  and dkey = 'PCNAME'"
        '        SQL = SQL & " and a.data =b.username "
        '        SQL = SQL & " order by a.unique_id"

        '        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        '        DPCNAME.Items.Add("")
        '        For i = 0 To dtReferp.Rows.Count - 1
        '            Dim ListItem1 As New ListItem
        '            ListItem1.Text = dtReferp.Rows(i).Item("Data")
        '            ListItem1.Value = dtReferp.Rows(i).Item("userid")
        '            If ListItem1.Value = pName Then ListItem1.Selected = True
        '            DPCNAME.Items.Add(ListItem1)
        '        Next

        '        dtReferp.Clear()
        '    End If
        'End If

        'If wStep > 110 Then
        '    '�ͺަ��b�H��
        '    If pFieldName = "PCNAME" Then
        '        DPCNAME.Items.Clear()

        '        SQL = " Select  a.*,userid  from M_referp a,M_users b"
        '        SQL = SQL & "   where  cat = '6001'"
        '        SQL = SQL & "  and dkey = 'PCNAME'"
        '        SQL = SQL & " and a.data =b.username "
        '        SQL = SQL & " AND b.userid ='" + pName + "'"
        '        SQL = SQL & " order by a.unique_id"

        '        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        '        DPCNAME.Items.Add("")
        '        For i = 0 To dtReferp.Rows.Count - 1
        '            Dim ListItem1 As New ListItem
        '            ListItem1.Text = dtReferp.Rows(i).Item("Data")
        '            ListItem1.Value = dtReferp.Rows(i).Item("userid")
        '            If ListItem1.Value = pName Then ListItem1.Selected = True
        '            DPCNAME.Items.Add(ListItem1)
        '        Next

        '        dtReferp.Clear()

        '    End If
        'End If




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
    '**     �x�s���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True And OK() Then


            DisabledButton()   '����Button�B�@
            FlowControl("OK", 0, "1")


        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG1���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            DisabledButton()   '����Button�B�@
            FlowControl("NG1", 1, "2")


        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG2���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then

            If wStep = 10 Then
                If OK() = True Then
                    DisabledButton()   '����Button�B�@
                    FlowControl("NG2", 2, "3")

                End If
            Else
                DisabledButton()   '����Button�B�@
                FlowControl("NG2", 2, "3")

            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     �y�{����
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   �U�@���d�ɨϥ� 
    '**        pSts=0:���B�z, 1:OK, 2:NG1, 3:NG2, 4:�w�\Ū, 5:�Q���  ��sT_Waithandle���A
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T


        GetDataStatus()  '���o�����d�U���s��Data Status



        'Check����z��
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--�ˬd�e�U��No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("006001", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '�O�_����
            Dim RepeatRun As Boolean = False    '�O�_���а���
            Dim MultiJob As Integer = 0         '�h�H�֩w
            Dim wLevel As String = ""           '�˫~������

            While Run = True
                Run = False     '����Flag=������
                '--���o�U�@���ѼƳ]�w---------
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=�q��
                Dim pCount As Integer
                '--���oLeadTime�ѼƳ]�w---------
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--���o�u�{�t���ѼƳ]�w---------
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--��L�ѼƳ]�w---------
                Dim i, RtnCode As Integer
                Dim NewFormSno As Integer = wFormSno    '���y����
                Dim pRunNextStep As Integer = 0         '�O�_����p��U�@��(�|ñ)
                Dim SQL As String = ""

                '���o���y�����Χ�s������
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '�P�_�O�_�_��
                        '���o���y����
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '��渹�X, ���y����
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else

                            '�ӽЬy�{��ƫظm
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo,Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                            'Modify-End
                            '�]�w�e�UNo
                            DNo.Text = SetNo(NewFormSno)

                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '���O�q�������а���
                            '��s������
                            ModifyTranData(pFun, pSts)

                            '�y�{��Ƶ���
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '��渹�X,���y����,�u�{���d���X,��ƾ�,ñ�֪�, �y�{�����_(�|ñ)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '�O�q�������а���
                        End If
                    End If
                End If

                '���o�U�@��
                If ErrCode = 0 And pRunNextStep = 1 Then

                    Dim wAllocateID As String = ""

                    If (wStep = 1 Or wStep = 500) And pFun = "OK" And DSTYPE.SelectedValue = "�릸" Then
                        wAllocateID = GetRelated(wApplyID)  '�x��
                        pAction = 0
                        If DDISPOSALTYPE.Value <> "DEPOT-07" Then
                            pAction = 0
                        Else
                            pAction = 2
                        End If

                    End If

                    If (wStep = 1 Or wStep = 502) And pFun = "OK" And DSTYPE.SelectedValue = "�g������" Then
                        pAction = 1
                    End If

                    If wStep = 6 Then
                        wAllocateID = GetRelated(wApplyID)  '�x��
                    End If



                    If wStep = 10 And pFun = "OK" Then
                        If wApplyID = "mf002-1" Then
                            wAllocateID = GetRelated("vf010-1")
                        ElseIf wApplyID = "mf002-2" Then
                            wAllocateID = GetRelated("vf010-2")
                        Else
                            wAllocateID = GetRelated(Request.QueryString("pUserID"))
                        End If


                    End If


                    '��Xñ�֪�User ID �x�ǥD�� 
                    If wStep = 111 Then  '���w�x�ǥD�� 
                        If pAction = 0 Then
                            wAllocateID = GetRelated(wApplyID)
                        End If
                    End If


                    '��Xñ�֪�User ID  
                    If wStep = 112 Then  '���w�ͺ޾���b
                        If pAction = 0 Then
                            wAllocateID = GetUserID(DPCNAME.Text)
                        End If
                    End If





                    '20240211 �������W�w
                    ''�p�G�O�ͺު��H�_��N���n�A�]�ͺ�ñ��
                    'If DDepoName.Text = "�Ͳ��޲z" And wStep = 20 And pFun = "OK" Then
                    '    If DMKTSIGN.Checked = True Then '��~
                    '        pAction = 2
                    '    Else
                    '        pAction = 3  '�d������
                    '        wAllocateID = DDUTYDEPOID.Text
                    '    End If
                    'End If

                    If wStep = 20 And pAction = 1 Then
                        wAllocateID = GetRelated(wApplyID)
                    End If

                    If wStep = 45 Then

                        If DMKTSIGN.Checked = True Then
                            '���oñ�֪� ID  
                            wAllocateID = getDUTYDEPOID(wFormSno)

                            '�W�[CC ���g�z
                            If DDISPOSALTYPE.Value = "PN" Or DDISPOSALTYPE.Value = "QF" Or DDISPOSALTYPE.Value = "S&B" Then
                                pAction = 3
                            Else
                                If DDUTYDEPO.SelectedValue = "��~����" Then
                                    pAction = 0
                                Else
                                    pAction = 2
                                End If

                                '�p�G����~ñ�֭n����50�d������
                            End If

                        Else
                            '���oñ�֪� ID  
                            '�p�G����~ñ�֭n����50�d������
                            wAllocateID = getDUTYDEPOID(wFormSno)


                            '�W�[CC ���g�z
                            If DDISPOSALTYPE.Value = "PN" Or DDISPOSALTYPE.Value = "QF" Or DDISPOSALTYPE.Value = "S&B" Then
                                pAction = 3
                            Else
                                If wAllocateID.ToUpper = "PC028" Then
                                    pAction = 1
                                Else
                                    pAction = 1
                                End If

                            End If


                        End If
                    End If

                    If wStep = 46 Then
                        wAllocateID = getDUTYDEPOID(wFormSno)
                    End If


                    If wStep = 50 And pFun = "OK" Then

                        If DMKTSIGN.Checked = True Then
                            wAllocateID = GetRelated(getDUTYDEPOID(wFormSno))

                        Else
                            wAllocateID = GetRelated(GetRelated(wApplyID)) '���X�������y�D��

                            If wAllocateID.ToUpper = "PC028" Then
                                pAction = 3
                            Else
                                pAction = 3
                            End If
                        End If

                    End If


                    '�p�G�S����~ñ�֭n����60�d������
                    If wStep = 50 And pFun = "NG1" Then
                        '���oñ�֪� ID

                        If DDepoName.Text = "�Ͳ��޲z" Then
                            pAction = 2
                            wAllocateID = GetRelated(wApplyID)
                        Else
                            pAction = 1

                        End If

                    End If

                    '20151223 �W�[CC��~���g�z 
                    If wStep = 51 And pFun = "OK" Then

                        If DMKTSIGN.Checked = True Then
                            wAllocateID = GetRelated(getDUTYDEPOID(wFormSno))
                            If wAllocateID = "mk002" Then
                                If DDUTYDEPO.SelectedValue <> "��~����" Then
                                    pAction = 0
                                Else
                                    pAction = 1
                                End If

                            End If
                        Else
                            '  wAllocateID = DDUTYDEPOID.Text
                            wAllocateID = getDUTYDEPOID(wFormSno)
                            If wAllocateID.ToUpper = "PC026" Then
                                wAllocateID = GetRelated(wAllocateID.ToUpper)
                            End If
                        End If


                    End If




                    ' If wStep = 30 And pFun = "NG1" Then
                    ' wAllocateID = GetRelated(GetRelated(wApplyID)) '���X�������y�D��
                    ' pAction = 1
                    'End If

                    If wStep = 60 And pFun = "OK" Then
                        If DMKTSIGN.Checked = True Then
                            pAction = 2
                        Else
                            pAction = 0
                        End If

                    End If




                    If wStep = 70 And pFun = "NG1" Then
                        pAction = 1
                        wAllocateID = GetRelated(GetRelated(wApplyID)) '�d�������D��1
                    End If


                    If wStep = 80 And pFun = "NG1" Then
                        wAllocateID = GetRelated(GetRelated(GetRelated(wApplyID))) '�d�������D��2
                        pAction = 1
                    End If



                    '�p�G������~ñ�֤SNG�n�����^��~
                    If wStep = 61 Then
                        wAllocateID = getDUTYDEPOID(wFormSno)
                    End If

                    If wStep = 62 Then
                        wAllocateID = GetRelated(Request.QueryString("pUserID"))
                    End If

                    If wStep = 31 Then
                        wAllocateID = GetRelated(wApplyID)
                    End If


                    If wStep = 82 Then
                        wAllocateID = GetRelated(GetRelated(wApplyID)) '���X�������y�D��
                    End If


             

                    '��X��ڳB�z�H�� ID  
                    If wStep = 115 Then   '���o��ڳB�z�H��
                        If pAction = 0 Then
                            wAllocateID = GetUserID(GetDEPMAN(DAppName.Text))
                        End If
                    End If




                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '��渹�X,�u�{���d���X,ñ�֪�,�ӽЪ�,�Q�N�z��,�Q���w��,�h�H�֩w�u�{No,
                    '�U�@�u�{��, ���X, ����, �Q�N�z��, �H��, �B�z��k, �ʧ@(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                    If ErrCode = 0 Then pAction = 0
                End If

                '�ظm�y�{���
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        wNextGate = ""
                        For i = 1 To pCount
                            '���o�U�@���H��(�T���ɨϥ�)
                            If wNextGate = "" Then
                                wNextGate = pNextGate(i)
                            Else
                                wNextGate = "," & pNextGate(i)
                            End If

                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo,Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '���o�֩w��-�s�զ�O��
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '���o�u�{�t���̫���
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '�֩w��, ��渹�X, �u�{���X, ���O(0:�q��,1:�֩w), �}�l���, �̫���, ���

                            '���o�w�w�}�l,������{�p��
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '��渹�X,�u�{���X,������,QC-L/T,�{�b�ɶ�, �w�w�}�l���, �w�w�������, ��ƾ�

                            '�ظm�y�{���
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,�ӽЪ�ID,��ƾ�,�ظm��, ñ�֪�, �Q�N�z��, �ӽЪ�, �w�w�}�l��, �w�w������, ���n��
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo,Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '��u�{��{�վ�
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        'ñ�֪�,��渹�X,���y����,�u�{���d���X,�Ǹ�,�{�b���,������,��ƾ�
                        'Modify-End
                    End If
                End If
                '�x�s�����
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '�P�_�O�_�_��
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If  'pSeqno <> 0
                    Else    '�P�_�O�_�_��
                        If pNextStep = 999 Then     '�u�{������?


                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '��s�����
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                        Else

                            ModifyData(pFun, "0")         '��s����� Sts=0(����)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    '�ǰe�l��
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            If pNextStep <> 30 And pNextStep <> 140 And pNextStep <> 130 And pNextStep <> 80 And wStep <> 70 And wStep <> 80 Then '�t�� �˥� �`�g�z�����H
                                oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                                '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O
                            End If
                            '    oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O
                        Next i
                    Else
                        oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O
                    End If

                    If pFlowType <> 3 Then
                        MultiJob = 0
                    Else
                        MultiJob = 1
                    End If

                    If (pRunNextStep = 1) And (pFlowType = 0 Or pFlowType = 3) Then
                        Run = True
                        RepeatRun = True
                        pAction = 0
                    End If



                    wStep = pNextStep     '�U�@�u�{���d���X

                    wFormSno = NewFormSno '�U�@�u�{���y����
                Else
                    EnabledButton()   '�_��Button�B�@
                    If ErrCode = 9110 Then Message = "���o���y�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9120 Then Message = "�y�{��Ƨ�s���`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9130 Then Message = "�U�u�{�p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9131 Then Message = "�L�U�u�{�޲z�H,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9140 Then Message = "�u�{�w�w�}�l�Χ�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9150 Then Message = "�U�@�u�{��ƫظm���`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9160 Then Message = "�u�{������ƫظm���`(999),�нT�{�γs���t�ΤH��!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '�x�s���ErrCode=0
            End While       '���а���


            If ErrCode = 0 Then
                '--�l��ǰe---------
                oCommon.SendMail()

                Dim URL As String



                URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                           "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID

                Response.Redirect(URL)


            End If



        Else
            EnabledButton()   '�_��Button�B�@
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
            If ErrCode = 9040 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            If ErrCode = 9050 Then Message = "���謰��L�ɻݶ�g����,�нT�{!"
            If ErrCode = 9060 Then Message = "�e�U��No.����,�нT�{�e�U��No.!"
            If ErrCode = 9070 Then Message = "�ݶ�J�u�{�ݳB�z�����Τu�{�ݳB�z��.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     ���o���i�ת��A
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            'NG-1���s
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            End If
            'NG-2���s
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            End If
            'OK���s
            If dtFlow.Rows(0)("OKFun") = 1 Then
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim FileName As String
        Dim SQl As String


        SQl = "Insert into F_DISPOSALSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno,DisposalYM,"
        SQl = SQl + "No,MKTSIGN,CUSTOMERTOLL,APPNAME,DEPONAME,"  '1~5         
        SQl = SQl + "APPDATE,DISPOSALREASON,DISPOSALREASON1,DUTYDEPO,Sales,PLACE,STYPE, "
        SQl = SQl + "DISPOSALRULE,DISPOSALTYPE,CHINESEREASON,JAPANREASON,DISPOSALFILE1,"
        SQl = SQl + "PCNAME,ASDate,checkASDate,SignDate,"
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '006001', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����
        If DSignDate.Text = "�L" Then
            SQl = SQl + " '" + Mid(Now.ToString("yyyy/MM/dd"), 1, 7) + "', "   '���o�~��
        Else

            SQl = SQl + " '" + Mid(DSignDate.Text, 1, 7) + "', "   '���o�~��
        End If

        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1

        If DMKTSIGN.Checked = True Then
            SQl = SQl + "1,"
        Else
            SQl = SQl + "0,"
        End If


        If DCUSTOMERTOLL.Checked = True Then
            SQl = SQl + "1,"
        Else
            SQl = SQl + "0,"
        End If

        SQl = SQl + " N'" + DAppName.Text + "', "                '�ӽЪ� 
        SQl = SQl + " N'" + DDepoName.Text + "', "                '����
        SQl = SQl + " N'" + DAppDate.Text + "', "                '���
        SQl = SQl + " N'" + DDISPOSALREASON.SelectedValue + "', "                '���o��]
        SQl = SQl + " N'" + DDISPOSALREASON1.Text + "', "                '���o��]
        SQl = SQl + " N'" + DDUTYDEPO.SelectedValue + "', "                '�d������
        SQl = SQl + " N'" + DSales.Text + "', "                '�~��
        SQl = SQl + " N'" + DPLACE.SelectedValue + "', "                '����
        SQl = SQl + " N'" + DSTYPE.SelectedValue + "', "                '����
        SQl = SQl + " N'" + DDISPOSALRULE.SelectedValue + "', "                '�ǫh
        SQl = SQl + " N'" + DDISPOSALTYPE.Value + "', "                '����
        SQl = SQl + " N'" + YKK.ReplaceString(DCHINESEREASON.Text) + "', "                '���o����
        SQl = SQl + " N'" + YKK.ReplaceString(DJAPANREASON.Text) + "', " '���o���
        '�m�W4
        FileName = ""
        If DFileUpload1.PostedFile.FileName <> "" Then
            Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                          CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
            Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                        System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")

            '20170912 �N�ɦW�ק令���t��l�ɦW
            Dim fileExtension As String  '���ɦW
            fileExtension = IO.Path.GetExtension(DFileUpload1.PostedFile.FileName).ToLower   '���o�ɮ׮榡

            FileName = CStr(NewFormSno) + "-DPFILE-" + UploadDateTime + fileExtension

            DFileUpload1.PostedFile.SaveAs(Path + FileName)

        Else
            FileName = ""
        End If

        SQl = SQl + "'" + FileName + "'," '���o����
        SQl = SQl + " N'" + DPCNAME.Text + "', "                '�ͺަ��b
        SQl = SQl + " '" + DASDate.Value + "', "                '�߷|�ɶ�
        '�߷|�ɶ�
        Dim j As Integer
        For j = 0 To CheckASDate.Items.Count - 1 Step j + 1
            ' check item selected state.
            If (CheckASDate.Items(j).Selected) Then
                SQl = SQl + CheckASDate.Items(j).Value + ","
                Exit For
            End If

        Next

        If DSTYPE.SelectedValue = "�릸" And DDISPOSALTYPE.Value <> "DEPOT-07" Then
            SQl = SQl + " '" + DSignDate.Text + "', "                'ñ�ִ���
        Else '�g�����ƹw�]�C��̫�@��
            SQl = SQl + "CONVERT(CHAR(10),dateadd(day,-day(getdate( )),dateadd(month,1,getdate())),111), "                'ñ�ִ���
        End If





        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "        '�@����
        SQl = SQl + " '" + NowDateTime + "', "                            '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                                     '�ק��
        SQl = SQl + " '" + NowDateTime + "' "                             '�ק�ɶ�
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)


        '�s�J���䴩����

        SQl = " delete from F_DISPOSALSheetdt"
        SQl = SQl + " where formsno ='" + CStr(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQl)




        Dim ITEMCODE, ITEMNAME, PIECE, METER, YARD, KG, PRICE, DISRULE, REMARK As String

        Dim sql2 As String
        Dim I As Integer
        For I = 1 To 10

            ITEMCODE = "DITEMCODE" + Trim(Str(I))
            Dim DITEMCODE As TextBox = Me.FindControl(ITEMCODE)

            ITEMNAME = "DITEMNAME" + Trim(Str(I))
            Dim DITEMNAME As TextBox = Me.FindControl(ITEMNAME)

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

            REMARK = "DREMARK" + Trim(Str(I))
            Dim DREMARK As TextBox = Me.FindControl(REMARK)


            If DITEMNAME.Text <> "" Then
                sql2 = " insert into  F_DISPOSALSheetdt (Sts,CompletedTime,Formno,FormSno,NO,SeqNO,ITEMCODE,ITEMNAME,"
                sql2 = sql2 + " PIECE,METER,YARD,KG,PRICE,DISRULE,REMARK,"
                sql2 = sql2 + " CreateUser, CreateTime, ModifyUser, ModifyTime)"
                sql2 = sql2 + " values('0',getdate(),'006001','" + CStr(NewFormSno) + "','" + DNo.Text + "','" + CStr(I) + "','" + YKK.ReplaceString(DITEMCODE.Text) + "','" + YKK.ReplaceString(DITEMNAME.Text) + "',"
                sql2 = sql2 + " '" + DPIECE.Text + "','" + DMETER.Text + "',"
                sql2 = sql2 + " '" + YKK.ReplaceString(DYARD.Text) + "','" + YKK.ReplaceString(DKG.Text) + "','" + YKK.ReplaceString(DPRICE.Text) + "','" + YKK.ReplaceString(DDISRULE.SelectedValue) + "','" + YKK.ReplaceString(DREMARK.Text) + "',"
                sql2 = sql2 + " '" + Request.QueryString("pUserID") + "', "        '�@����
                sql2 = sql2 + " '" + NowDateTime + "', "                            '�@���ɶ�
                sql2 = sql2 + " '" + "" + "', "                                     '�ק��
                sql2 = sql2 + " '" + NowDateTime + "' "                             '�ק�ɶ�
                sql2 = sql2 + " ) "
                uDataBase.ExecuteNonQuery(sql2)
            End If


        Next




    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     ��s�����
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)

        Dim SQL As String = ""

        '�D�ɸ��
        SQL = " Update   F_DISPOSALSheet"
        SQL = SQL + " Set "
        If pFun <> "SAVE" Then
            SQL = SQL + " Sts = '" & pSts & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
        End If
        If DMKTSIGN.Checked = True Then
            SQL = SQL + "MKTSIGN =1,"
        Else
            SQL = SQL + "MKTSIGN =0,"
        End If


        If DCUSTOMERTOLL.Checked = True Then
            SQL = SQL + "CUSTOMERTOLL=1,"
        Else
            SQL = SQL + "CUSTOMERTOLL=0,"
        End If
        SQL = SQL + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " APPDate = N'" & DAppDate.Text & "',"
        SQL = SQL + " DepoName = N'" & DDepoName.Text & "',"
        SQL = SQL + " APPName = N'" & DAppName.Text & "',"
        SQL = SQL + " DISPOSALREASON= N'" & DDISPOSALREASON.SelectedValue & "',"
        SQL = SQL + " DISPOSALREASON1= N'" & DDISPOSALREASON1.Text & "',"
        SQL = SQL + " DUTYDEPO= N'" & DDUTYDEPO.SelectedValue & "',"
        SQL = SQL + " SALES= N'" & DSales.Text & "',"
        SQL = SQL + " DISPOSALTYPE= N'" & DDISPOSALTYPE.Value & "',"
        SQL = SQL + " PLACE=N'" & DPLACE.SelectedValue & "',"
        SQL = SQL + " STYPE=N'" & DSTYPE.SelectedValue & "',"
        SQL = SQL + " DISPOSALRULE= N'" & DDISPOSALRULE.SelectedValue & "',"
        SQL = SQL + " CHINESEREASON= N'" & YKK.ReplaceString(DCHINESEREASON.Text) & "',"
        SQL = SQL + " JAPANREASON= N'" & YKK.ReplaceString(DJAPANREASON.Text) & "',"
        Dim FileName As String

        FileName = ""
        If DFileUpload1.Visible = True Then
            '   And LCertifcateFile.NavigateUrl = "" Then
            If DFileUpload1.PostedFile.FileName <> "" Or LDISPOSALFILE.NavigateUrl <> "" Then
                Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                            System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")

                '20170912 �N�ɦW�ק令���t��l�ɦW
                Dim fileExtension As String  '���ɦW
                fileExtension = IO.Path.GetExtension(DFileUpload1.PostedFile.FileName).ToLower   '���o�ɮ׮榡

                FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime + fileExtension

                If DFileUpload1.PostedFile.FileName = "" Then
                    '  FileName = Right(LDISPOSALFILE.NavigateUrl, InStr(StrReverse(LDISPOSALFILE.NavigateUrl), "\") - 1)
                    DFileUpload1.PostedFile.SaveAs(Path + FileName)
                Else
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime & "-" & Right(DFileUpload1.PostedFile.FileName, InStr(StrReverse(DFileUpload1.PostedFile.FileName), "\") - 1)
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime + fileExtension
                    DFileUpload1.PostedFile.SaveAs(Path + FileName)
                End If

            Else
                FileName = ""
            End If
            SQL = SQL + " DISPOSALFILE1='" + FileName + "'," '���o�ҩ�
        End If
        SQL = SQL + " PIECE ='" + DPIECE13.Text + "',"
        SQL = SQL + " METER ='" + DMETER13.Text + "',"
        SQL = SQL + " YARD ='" + DYARD13.Text + "',"
        SQL = SQL + " KG ='" + DKG13.Text + "',"
        SQL = SQL + " PRICE ='" + DPRICE13.Text + "',"
        SQL = SQL + " PCNAME= N'" & DPCNAME.Text & "',"
        SQL = SQL + " ASDate= '" & DASDate.Value & "',"
        SQL = SQL + " SignDate= '" & DSignDate.Text & "',"



        '�߷|�ɶ�
        Dim j As Integer
        For j = 0 To CheckASDate.Items.Count - 1 Step j + 1
            ' check item selected state.
            If (CheckASDate.Items(j).Selected) Then
                SQL = SQL + "CheckASDate=" + CheckASDate.Items(j).Value + ","
                Exit For
            End If

        Next


        SQL = SQL + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
        SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(SQL)

        '���Ӹ��



        '�s�J���䴩����


        SQL = " delete from F_DISPOSALSheetdt"
        SQL = SQL + " where no ='" + DNo.Text + "'"
        uDataBase.ExecuteNonQuery(SQL)


        Dim ITEMCODE, ITEMNAME, PIECE, METER, YARD, KG, PRICE, DISRULE, REMARK As String

        Dim sql2 As String
        Dim I As Integer
        For I = 1 To 10

            ITEMCODE = "DITEMCODE" + Trim(Str(I))
            Dim DITEMCODE As TextBox = Me.FindControl(ITEMCODE)

            ITEMNAME = "DITEMNAME" + Trim(Str(I))
            Dim DITEMNAME As TextBox = Me.FindControl(ITEMNAME)

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

            REMARK = "DREMARK" + Trim(Str(I))
            Dim DREMARK As TextBox = Me.FindControl(REMARK)


            If DITEMNAME.Text <> "" Then
                sql2 = " insert into  F_DISPOSALSheetdt (Sts,CompletedTime,Formno,FormSno,NO,SeqNO,ITEMCODE,ITEMNAME,"
                sql2 = sql2 + " PIECE,METER,YARD,KG,PRICE,DISRULE,REMARK,"
                sql2 = sql2 + " CreateUser, CreateTime, ModifyUser, ModifyTime)"
                sql2 = sql2 + " values('0',getdate(),'006001','" + CStr(wFormSno) + "','" + DNo.Text + "','" + CStr(I) + "','" + YKK.ReplaceString(DITEMCODE.Text) + "','" + YKK.ReplaceString(DITEMNAME.Text) + "',"
                sql2 = sql2 + " '" + DPIECE.Text + "','" + DMETER.Text + "',"
                sql2 = sql2 + " '" + YKK.ReplaceString(DYARD.Text) + "','" + YKK.ReplaceString(DKG.Text) + "','" + YKK.ReplaceString(DPRICE.Text) + "','" + YKK.ReplaceString(DDISRULE.SelectedValue) + "','" + YKK.ReplaceString(DREMARK.Text) + "',"
                sql2 = sql2 + " '" + Request.QueryString("pUserID") + "', "        '�@����
                sql2 = sql2 + " '" + NowDateTime + "', "                            '�@���ɶ�
                sql2 = sql2 + " '" + "" + "', "                                     '�ק��
                sql2 = sql2 + " '" + NowDateTime + "' "                             '�ק�ɶ�
                sql2 = sql2 + " ) "
                uDataBase.ExecuteNonQuery(sql2)
            End If


        Next



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     ��s������
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim SQl As String



        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"
            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        uDataBase.ExecuteNonQuery(SQl)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     �l�[�����ƩM�e�U���Ӫ�
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String


        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQl)

        If DBAdapter1.Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    uDataBase.ExecuteNonQuery(SQl)
                End If
            End If
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '�w�q���\���ɮ׮榡
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

        'If UPFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
        'Check�W���ɮ׮榡
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����Button(OK, Ng1, NG2, Save)�B�@ 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �_��Button(OK, Ng1, NG2, Save)�B�@ 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    �]�w���s,����,����,�֩w�i������m
    '**
    '*****************************************************************
    Sub SetControlPosition()



        If DDescSheet.Visible Then                                      ' ����
            DDescSheet.Style("top") = Top - 140 & "px"
            DDecideDesc.Style("top") = Top - 140 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' ���𻡩�
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top + 20 & "px"
        BNG1.Style("top") = Top + 20 & "px"
        BNG2.Style("top") = Top + 20 & "px"
        BOK.Style("top") = Top + 20 & "px"


        Top = Top + 50
        If GridView2.Rows.Count > 0 Then                                ' �֩w�i��
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
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
        Str = "D" + CStr(DateTime.Now.Year) + Str2
        '�~
        'Set�渹
        '�������Ʀ��X��  20150414 Modify by Jessica
        Dim sql As String
        sql = " select  isnull(right(max(no),3),0) as  cun from  F_DISPOSALSheet  "
        sql = sql + " where  left(convert(char(10),appdate ,111),7) = left(convert(char(10), getdate(),111),7)"
        Dim dt1 As DataTable = uDataBase.GetDataTable(sql)
        If dt1.Rows.Count > 0 Then
            Str1 = CStr(CInt(dt1.Rows(0).Item("cun")) + 1)
        End If

        For i = Str1.Length To 3 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��J�ˬd
    '**
    '*****************************************************************
    Function OK() As Boolean
        If wStep = 1 Or wStep = 500 Then

            Top = 1200
        Else
            Top = 1400
        End If

        SetControlPosition()
        '  ShowFormData()
        If wStep = 1 Then
            DDISPOSALREASON1.BackColor = Color.Yellow
            DSales.BackColor = Color.Yellow
        End If


        '�P�_�O�_���i�ӽг��o

        If wStep = 1 Then
            If DSTYPE.Text = "�릸" Then
                If checkDate() = 0 Then
                    isOK = False
                    Message = "���`�G" + "�w�L�ӽд���:" + SignDate + ",�U���}��ӽФ��:" + OpenDate

                End If
            End If

        End If

        '20241231

        If DSignDate.Text = "���}��" Then
            Message = "���`�Gñ�ִ���=���}�񤣯�ӽ�!"
            isOK = False
        End If




        '�p�G��䥦�n�����z��



        Dim I As Integer
        Dim ITEMNAME, PIECE, METER, YARD, KG, PRICE, DISRULE, wRow As String

        For I = 1 To 10

            wRow = "��" + CStr(I) + "�C"

            ITEMNAME = "DITEMNAME" + Trim(Str(I))
            Dim DITEMNAME As TextBox = Me.FindControl(ITEMNAME)

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


            '�ˬd �ƶq�_�����
            If I = 1 Then
                If DITEMNAME.Text = "" Then
                    isOK = False
                    Message = "���`�G�ݿ�JITEMNAME!"
                    Exit For
                End If

            End If


            '�ˬd ITEMNAME �O�_�����
            If DDISRULE.SelectedValue <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DKG.Text <> "" Or DPRICE.Text <> "" Or DYARD.Text <> "" Then
                If DITEMNAME.Text = "" Then
                    isOK = False
                    Message = "���`�G�ݿ�JITEMNAME!"
                End If
            End If

            If DDISRULE.SelectedValue <> "" Or DITEMNAME.Text <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DPRICE.Text <> "" Or DYARD.Text <> "" Then
                If DKG.Text = "" Then '�ˬd ���q�_�����
                    isOK = False
                    Message = Message + "\n" + "���`�G�ݿ�J���q!"
                End If
            End If

            If DDISRULE.SelectedValue <> "" Or DITEMNAME.Text <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DKG.Text <> "" Or DYARD.Text <> "" Then
                If DPRICE.Text = "" Then '�ˬd ���B�O�_�����
                    isOK = False
                    Message = Message + "\n" + "���`�G�ݿ�J���B!"
                End If
            End If

            If DITEMNAME.Text <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DKG.Text <> "" Or DPRICE.Text <> "" Or DYARD.Text <> "" Then
                If DDISRULE.SelectedValue = "" Then '�ˬd ���o�ǫh�O�_�����
                    isOK = False
                    Message = Message + "\n" + "���`�G�ݿ�J���o�ǫh!"
                End If
            End If


            '  If DITEMNAME.Text <> "" Or DKG.Text <> "" Or DPRICE.Text <> "" Or DDISRULE.SelectedValue <> "" Then
            'If DPIECE.Text = "" And DMETER.Text = "" And DYARD.Text = "" Then
            ' isOK = False
            ' Message = Message + "\n" + "���`�G�ݿ�J�ƶq�T��@!"
            ' End If
            ' End If



            If Message <> "" Then
                If checkDate() <> 0 Then
                    Message = wRow + "\n" + Message
                End If

                Exit For
            End If

        Next


        If DDISPOSALREASON.SelectedValue = "�䥦" Then
            If DDISPOSALREASON1.Text = "" Then
                isOK = False
                Message = Message + "\n" + "���`�G�ӽЭ�]��䥦�A�ݿ�J����!"
            End If
            If wStep = 1 Then
                DDISPOSALREASON1.BackColor = Color.GreenYellow
            End If

        End If


        If DDUTYDEPO.SelectedValue = "��~" Then
            If DSales.Text = "" Then
                isOK = False
                Message = Message + "\n" + "���`�G�d�������O��~�A�ݿ�J�~�ȩm�W!"
            End If
            If wStep = 1 Then
                DSales.BackColor = Color.GreenYellow
            End If

        End If

        If DSTYPE.SelectedValue = "�릸" And DDISPOSALTYPE.Value <> "DEPOT-07" Then
            If DPCNAME.Text = "" Then
                isOK = False
                Message = Message + "\n" + "���`�G���ӽЪ̵L�]�w���b�H��" + "\n" + "�A�лP�ͺ޳��o���H���T�{!"
            End If

        End If

        '�P�_�O�_�ݥ߷|
        If wStep = 45 Then
            Dim j As Integer
            Dim checkflag As Integer = 0
            For j = 0 To CheckASDate.Items.Count - 1 Step j + 1
                ' check item selected state.
                If (CheckASDate.Items(j).Selected) Then
                    checkflag = 1
                    Exit For
                End If

            Next
            If checkflag = 0 Then
                isOK = False
                Message = Message + "\n" + "���`�G�ФĿ�O�_�ݥ߷|!"
            End If
        End If

        'If wStep = 110 Then
        '    Dim j As Integer
        '    For j = 0 To CheckASDate.Items.Count - 1 Step j + 1
        '        ' check item selected state.
        '        If (CheckASDate.Items(j).Selected) Then
        '            If CheckASDate.Items(j).Value = 1 Then
        '                DASDate.Style.Add("background-color", "greenyellow")
        '                If DASDate.Value = "" Then
        '                    isOK = False
        '                    Message = Message + "\n" + "���`�G�п�J�߷|�ɶ�!"
        '                End If

        '            End If
        '            Exit For
        '        End If

        '    Next
        'End If

        griderror = True
        inserterror = True
        fielderror = True

        If isOK = True Then
            If wStep = 1 Or wStep = 500 Then
                upload()
                If griderror = False Then
                    Message = Message + "\n" + "�W�Ǯ榡���~gridview,�нT�{!"
                    isOK = False
                End If
                If fielderror = False Then
                    Message = Message + "\n" + "CODE �� ���o�ǫh �� ���� �����\�ť�!"
                    isOK = False
                End If
                If isOK = True Then
                    Insert()
                    If inserterror = False Then
                        Message = Message + "\n" + "�W�Ǯ榡���~Insert,�нT�{!"
                        isOK = False
                    End If
                End If

            End If
        End If



        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If



        Return isOK


    End Function

    Sub getDUTYDEPOID()
        Dim sql As String

        If Trim(DDUTYDEPO.SelectedValue) <> "�u�t" Then

            If DDISPOSALTYPE.Value = "PN" Or DDISPOSALTYPE.Value = "QF" Or DDISPOSALTYPE.Value = "S&B" Then
                sql = " select  userid,username  from M_users "
                sql = sql + " where username in ("
                sql = sql + "  select data from m_referp"
                sql = sql + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(DDISPOSALTYPE.Value) + "'"
                sql = sql + " )"
            Else
                sql = " select  userid,username  from M_users "
                sql = sql + " where username in ("
                sql = sql + "  select data from m_referp"
                sql = sql + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(DDUTYDEPO.SelectedValue) + "'"
                sql = sql + " )"
            End If

            Dim dt2 As DataTable = uDataBase.GetDataTable(sql)
            If dt2.Rows.Count > 0 Then
                DDUTYDEPOID.Text = dt2.Rows(0).Item("userid")
                DDUTYDEPOName.Text = dt2.Rows(0).Item("username")
                ' DDUTYDEPOID.Text = "it013"
                'DDUTYDEPOName.Text = "�i�Q��"
            End If
        Else

            If DMKTSIGN.Checked Then  ' ����~ñ��
                If DDISPOSALTYPE.Value = "PN" Or DDISPOSALTYPE.Value = "QF" Or DDISPOSALTYPE.Value = "S&B" Then
                    sql = " select  userid,username  from M_users "
                    sql = sql + " where username in ("
                    sql = sql + "  select data from m_referp"
                    sql = sql + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(DDISPOSALTYPE.Value) + "'"
                    sql = sql + " )"
                Else
                    sql = " select  userid,username  from M_users "
                    sql = sql + " where username in ("
                    sql = sql + "  select data from m_referp"
                    sql = sql + " where cat = 6001 and dkey = 'DUTYDEPO-��~'"
                    sql = sql + " )"
                End If

                Dim dt2 As DataTable = uDataBase.GetDataTable(sql)
                If dt2.Rows.Count > 0 Then
                    DDUTYDEPOID.Text = dt2.Rows(0).Item("userid")
                    DDUTYDEPOName.Text = dt2.Rows(0).Item("username")
                    '  DDUTYDEPOID.Text = "it013"
                    '  DDUTYDEPOName.Text = "�i�Q��"
                End If
            Else
                sql = " select ruserid,rusername  from m_condition"
                sql = sql + " where userid in ( select ruserid from m_condition"
                sql = sql + " where userid = '" + wApplyID + "'"
                sql = sql + " and relatedid = 'd')"
                sql = sql + " and relatedid = 'd'"
                Dim dt1 As DataTable = uDataBase.GetDataTable(sql)
                If dt1.Rows.Count > 0 Then
                    DDUTYDEPOID.Text = dt1.Rows(0).Item("ruserid")
                    DDUTYDEPOName.Text = dt1.Rows(0).Item("rusername")
                    '  DDUTYDEPOID.Text = "it013"
                    '   DDUTYDEPOName.Text = "�i�Q��"
                End If
            End If
         

        End If
    End Sub

    Protected Sub DDISPOSALREASON_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDISPOSALREASON.SelectedIndexChanged
        DCUSTOMERTOLL.Checked = False
        If DDISPOSALREASON.SelectedValue = "11.��L" Then
            DDISPOSALREASON1.BackColor = Color.GreenYellow
        Else
            DDISPOSALREASON1.BackColor = Color.Yellow
        End If
        If DDUTYDEPO.SelectedValue = "��~" Or DDUTYDEPO.SelectedValue = "��~����" Then
            DSales.BackColor = Color.GreenYellow
            If DDUTYDEPO.SelectedValue = "��~����" Then
                DCUSTOMERTOLL.Checked = True
            End If
        Else
            DSales.BackColor = Color.Yellow
        End If
        If DDISPOSALREASON.SelectedValue = "02.�w�ʫ~�G2�~�H�W���ϥ�" Or DDISPOSALREASON.SelectedValue = "05.2�~�H�WSLD�󽦫~" Or DDISPOSALREASON.SelectedValue = "06.�b�w�j��3�~�ϥζq(�L��)" Or DDISPOSALREASON.SelectedValue = "12. �w�ʫ~�G1.5�~�H�W���ϥ�" Then
            DDISPOSALRULE.SelectedValue = "�C���k��H"
        Else
            DDISPOSALRULE.SelectedValue = ""
        End If

        If DDISPOSALREASON.SelectedValue = "13.����/�j�f�Ѿl(FA)" Then
            DSTYPE.SelectedValue = "�g������"
        Else
            DSTYPE.SelectedValue = "�릸"
        End If

        If wStep = 1 Then  '�ͺޥD��KEY�߷|���

            Dim sql As String
            sql = "SELECT  sdate,edate, convert(char(10),dateadd(day,-1,convert(datetime, sdate, 111)),111)sdate1,convert(char(10),dateadd(day ,-1, dateadd(m, datediff(m,0,getdate())+1,0)),111) lastDate "
            sql = sql + " FROM ("
            sql = sql + " SELECT  SUBSTRING(DKEY,13,1) AS  DKEYS, DATA AS SDATE   from M_referp"
            sql = sql + " where(cat = 6001)"
            sql = sql + " and left(dkey,12)='SendTime-PCS'"
            sql = sql + " )A, ("
            sql = sql + " SELECT SUBSTRING(DKEY,13,1) AS  DKEYE,DATA AS EDATE   from M_referp"
            sql = sql + " where(cat = 6001)"
            sql = sql + " and left(dkey,12)='SendTime-PCE'"
            sql = sql + " )B "
            sql = sql + " WHERE(A.DKEYS = B.DKEYE)"
            sql = sql + " and left(sdate,7) =left(convert(char(10),getdate(),111),7) and sdate >= convert(char(10),getdate(),111) "
            sql = sql + " order by sdate "
            Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)


            If dtFlow.Rows.Count > 0 Then

                If DSTYPE.SelectedValue = "�g������" Then
                    DSignDate.Text = dtFlow.Rows(0)("lastdate")
                Else

                    DSignDate.Text = dtFlow.Rows(0)("lastdate")

                End If
            Else
                DSignDate.Text = "���}��"

            End If
        End If


        DDISRULEChang()

    End Sub

    Sub GetTotal()
        Dim I As Integer
        Dim SUMPIECE1, SUMPIECE2, SUMMETER1, SUMMETER2, SUMYARD1, SUMYARD2, SUMKG1, SUMKG2, SUMPRICE1, SUMPRICE2 As Double
        Dim PIECE, METER, YARD, KG, PRICE, DISRULE, REMARK As String

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


            REMARK = "DREMARK" + Trim(Str(I))
            Dim DREMARK As TextBox = Me.FindControl(REMARK)

            '20180815 monica 
            If DDISPOSALTYPE.Value = "�����~" Then
                If DPRICE.Text <> "" Then
                    If CInt(DPRICE.Text) > 100000 Then
                        DREMARK.Text = "�X�p"
                    Else
                        DREMARK.Text = ""
                    End If
                End If

            End If



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

        DPRICE11.Text = SUMPRICE1.ToString("#,0")
        DPRICE11.Style.Add("TEXT-ALIGN", "right")
        DPRICE12.Text = SUMPRICE2.ToString("#,0")
        DPRICE12.Style.Add("TEXT-ALIGN", "right")
        DPIECE13.Text = (SUMPIECE1 + SUMPIECE2).ToString("#,0.00")
        DPRICE13.Style.Add("TEXT-ALIGN", "right")
        DMETER13.Text = (SUMMETER1 + SUMMETER2).ToString("#,0.00")
        DMETER13.Style.Add("TEXT-ALIGN", "right")
        DYARD13.Text = (SUMYARD1 + SUMYARD2).ToString("#,0.00")
        DYARD13.Style.Add("TEXT-ALIGN", "right")
        DKG13.Text = (SUMKG1 + SUMKG2).ToString("#,0.00")
        DKG13.Style.Add("TEXT-ALIGN", "right")
        DPRICE13.Text = (SUMPRICE1 + SUMPRICE2).ToString("#,0")
        DPIECE13.Style.Add("TEXT-ALIGN", "right")

    End Sub

    Protected Sub DSUM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSUM.Click
        GetTotal()
    End Sub

  
    Protected Sub DDISPOSALFILE2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDISPOSALFILE2.Click
     

    End Sub
    Function checkDate() As Integer
        '�P�_�O�_�i�ӽг��o
        Dim SQL, SQL1, Today As String
        Dim SFlag As Integer = 0
        SignDate = ""
        OpenDate = ""
        Today = Now.ToString("yyyy/MM/dd") '�{�b���
        ' If wStep = 1 Then
        'Today = Now.ToString("yyyy/MM/dd") '�{�b���
        ' Today = "2015/12/11"
        'Else
        'Today = DSignDate.Text
        'End If


        SQL = "SELECT  sdate,edate, convert(char(10),dateadd(day,-1,convert(datetime, sdate, 111)),111)sdate1,convert(char(10),dateadd(day ,-1, dateadd(m, datediff(m,0,getdate())+1,0)),111)lastdate  "
        SQL = SQL + " FROM ("
        SQL = SQL + " SELECT  SUBSTRING(DKEY,13,1) AS  DKEYS, DATA AS SDATE   from M_referp"
        SQL = SQL + " where(cat = 6001)"
        SQL = SQL + " and left(dkey,12)='SendTime-PCS'"
        SQL = SQL + " )A, ("
        SQL = SQL + " SELECT SUBSTRING(DKEY,13,1) AS  DKEYE,DATA AS EDATE   from M_referp"
        SQL = SQL + " where(cat = 6001)"
        SQL = SQL + " and left(dkey,12)='SendTime-PCE'"
        SQL = SQL + " )B "
        SQL = SQL + " WHERE(A.DKEYS = B.DKEYE)"
        If wStep = 1 Then
            SQL = SQL + " and  '" + Today + "' between edate and sdate "
        Else
            SQL = SQL + " and  left(sdate,7) ='" + DDisposalYM.Text + "'"
        End If


        SQL1 = SQL + " order by sdate "

        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)

        Dim I As Integer
        For I = 0 To dtFlow.Rows.Count - 1
            If dtFlow.Rows(0)("Sdate") >= Today Then
                'If wStep = 1 Then
                SignDate = dtFlow.Rows(0)("lastdate")
                SFlag = 1
                'End If

                Exit For

                DSignDate.Text = SignDate

            Else
                '�U���}���
                SignDate = dtFlow.Rows(0)("sdate")
                OpenDate = dtFlow.Rows(0)("Edate")

                'Dim dtFlow2 As DataTable = uDataBase.GetDataTable(SQL3)
                'For K = 0 To dtFlow2.Rows.Count - 1
                ' If dtFlow2.Rows(0)("Sdate") >= Today Then

                'End If
                'Next
                SFlag = 0
                Exit For
            End If
        Next

        Return SFlag



    End Function
   
    Protected Sub DMKTSIGN_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMKTSIGN.CheckedChanged
        If DMKTSIGN.Checked Then
            getDUTYDEPOID()
        End If
    End Sub


    '���o���Y�H
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID , RUserName  from M_Related where userid='" & userId & "' and RelatedID='D'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("RUserID")
        End If
        Return NextGate
    End Function

    '���s���o�d���������Y�H
    Function GETDUTYDEPOID(ByVal value As String) As String
        Dim UserId As String = ""
        Dim SQL As String
        SQL = " select * from f_disposalsheet"
        SQL = SQL + " where formsno ='" + value.ToString + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
        If Trim(dt.Rows(0)("DutyDepo")) <> "�u�t" And Trim(dt.Rows(0)("DutyDepo")) <> "�T��" Then

            If Trim(dt.Rows(0)("DISPOSALTYPE")) = "PN" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "QF" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "S&B" Then
                SQL = " select  userid,username  from M_users "
                SQL = SQL + " where username in ("
                SQL = SQL + "  select data from m_referp"
                SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(dt.Rows(0)("DISPOSALTYPE")) + "'"
                SQL = SQL + " )"
            Else
                SQL = " select  userid,username  from M_users "
                SQL = SQL + " where username in ("
                SQL = SQL + "  select data from m_referp"
                SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(dt.Rows(0)("DUTYDEPO")) + "'"
                SQL = SQL + " )"
            End If

            Dim dt2 As DataTable = uDataBase.GetDataTable(SQL)
            If dt2.Rows.Count > 0 Then
                UserId = dt2.Rows(0).Item("userid")
            End If
        Else

            If Trim(dt.Rows(0)("MKTSIGN")) = 1 Then  ' ����~ñ��
                If Trim(dt.Rows(0)("DISPOSALTYPE")) = "PN" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "QF" Or Trim(dt.Rows(0)("DISPOSALTYPE")) = "S&B" Then
                    SQL = " select  userid,username  from M_users "
                    SQL = SQL + " where username in ("
                    SQL = SQL + "  select data from m_referp"
                    SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-" + Trim(dt.Rows(0)("DISPOSALTYPE")) + "'"
                    SQL = SQL + " )"
                Else
                    SQL = " select  userid,username  from M_users "
                    SQL = SQL + " where username in ("
                    SQL = SQL + "  select data from m_referp"
                    SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-��~'"
                    SQL = SQL + " )"
                End If

                Dim dt2 As DataTable = uDataBase.GetDataTable(SQL)
                If dt2.Rows.Count > 0 Then
                    UserId = dt2.Rows(0).Item("userid")
                End If
            Else

                SQL = " select ruserid,rusername  from m_condition"
                SQL = SQL + " where userid in ( select ruserid from m_condition"
                SQL = SQL + " where userid = '" + Trim(dt.Rows(0)("Createuser")) + "'"
                SQL = SQL + " and relatedid = 'd')"
                SQL = SQL + " and relatedid = 'd'"
               
                Dim dt1 As DataTable = uDataBase.GetDataTable(SQL)
                If dt1.Rows.Count > 0 Then
                    UserId = dt1.Rows(0).Item("ruserid")
                End If
              
            End If


        End If

        ' UserId = "it013"
        Return UserId
    End Function

  
    Protected Sub DDUTYDEPO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDUTYDEPO.TextChanged
        DCUSTOMERTOLL.Checked = False
        If DDUTYDEPO.SelectedValue = "��~" Or DDUTYDEPO.SelectedValue = "��~����" Then
            DMKTSIGN.Checked = True
            DSales.BackColor = Color.GreenYellow

            If DDUTYDEPO.SelectedValue = "��~����" Then
                DCUSTOMERTOLL.Checked = True
            End If

        Else
            DMKTSIGN.Checked = False
            DSales.BackColor = Color.Yellow
        End If
        If DDISPOSALREASON.SelectedValue = "11.��L" Then

            DDISPOSALREASON1.BackColor = Color.GreenYellow
        Else
            DDISPOSALREASON1.BackColor = Color.Yellow
        End If

        uJavaScript.PopMsg(Me, "�ЦA���T�{�d�������B�O�_�ݤĿ���~ñ��/�V�Ȥ����")

    End Sub

    Protected Sub DDISPOSALRULE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDISPOSALRULE.SelectedIndexChanged


        Dim Rule As String
        Dim i As Integer
        Dim sql As String

        For i = 1 To 10
            Rule = "DDISRULE" + Trim(Str(i))
            Dim DText As DropDownList = Me.FindControl(Rule)
            DText.Items.Clear()
        Next

        If DDISPOSALRULE.SelectedValue = "�C���k��H" Then
            '�Ӷ����o�ǫhappend

            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '6001'"
            sql = sql & " and dkey = 'RULE'"
            sql = sql & " and data = '�C���k'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

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
                    '   DText.SelectedValue = dtr("Data")
                Next

            Next

        Else

            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '6001'"
            sql = sql & " and dkey = 'RULE'"
            sql = sql & " and data = '�D�C���k'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

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
                    ' DText.SelectedValue = dtr("Data")
                Next

            Next

        End If

        If DITEMNAME1.Text <> "" Then
            DDISRULE1.SelectedIndex = 1
        Else
            DDISRULE1.SelectedIndex = 0
        End If


        If DITEMNAME2.Text <> "" Then
            DDISRULE2.SelectedIndex = 1
        Else
            DDISRULE2.SelectedIndex = 0
        End If


        If DITEMNAME3.Text <> "" Then
            DDISRULE3.SelectedIndex = 1
        Else
            DDISRULE3.SelectedIndex = 0
        End If


        If DITEMNAME4.Text <> "" Then
            DDISRULE4.SelectedIndex = 1
        Else
            DDISRULE4.SelectedIndex = 0
        End If


        If DITEMNAME5.Text <> "" Then
            DDISRULE5.SelectedIndex = 1
        Else
            DDISRULE5.SelectedIndex = 0
        End If


        If DITEMNAME6.Text <> "" Then
            DDISRULE6.SelectedIndex = 1
        Else
            DDISRULE6.SelectedIndex = 0
        End If


        If DITEMNAME7.Text <> "" Then
            DDISRULE7.SelectedIndex = 1
        Else
            DDISRULE7.SelectedIndex = 0
        End If


        If DITEMNAME8.Text <> "" Then
            DDISRULE8.SelectedIndex = 1
        Else
            DDISRULE8.SelectedIndex = 0
        End If


        If DITEMNAME9.Text <> "" Then
            DDISRULE9.SelectedIndex = 1
        Else
            DDISRULE9.SelectedIndex = 0
        End If


        If DITEMNAME10.Text <> "" Then
            DDISRULE10.SelectedIndex = 1
        Else
            DDISRULE10.SelectedIndex = 0
        End If



    End Sub


    Function upload() As Boolean
        ' Try
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
        '  Catch ex As Exception
        'uJavaScript.PopMsg(Me, "�W���ɮ׮榡���~upload,�нT�{!")
        'griderror = False
        'End Try

        
    End Function

    Sub Insert()
        '�ˬd�O�_���P����P�ӽФH���ɮסA�Y���h�N���e�����R��
        Dim a As String = ""
        Dim uploadflag As Integer = 1
        Dim nullflag As Integer = 1
        Dim sno As String
        Dim NewFormSno As Integer = wFormSno    '���y����
        Dim SQL As String
        Dim smonth As String
        If wStep = 1 Then
            sno = SetNo(NewFormSno)
        Else
            sno = DNo.Text
        End If

        SQL = " select  distinct  appname,smonth   from f_disposaldata"
        'SQL = SQL + "  where appname = '" + DAppName.Text + "'"
        SQL = SQL + " WHERE  No = '" + sno + "'"
        Dim dt1 As DataTable = uDataBase.GetDataTable(SQL)
        If dt1.Rows.Count > 0 Then
          
            SQL = " delete  from f_disposaldata"
            SQL = SQL + "  where no= '" + sno + "'"
            uDataBase.ExecuteNonQuery(SQL)
            uploadflag = 1
       

        End If
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID

        smonth = Mid(DAppDate.Text, 1, 7)


        ' If uploadflag = 1 Then
        Try

            '�W�Ǩ��Ʈw
            Dim j As Integer
            Dim jSQL As String
            jSQL = ""

            Dim i As Integer


            For i = 0 To Me.GridView1.Rows.Count - 1 Step i + 1



                SQL = "Insert into F_DISPOSALData (APPDATE,APPNAME,APPDEPO,SMONTH,No,Seqno,Code,ITEMNAME1,ITEMNAME2,LENGTH,U,Color,LOCATION,ACTUAL,FREE,SIZE,CHAIN,CLS,UNIT,UNITWEIGHT,"
                SQL = SQL + "WEIGHTKG,DisposalRule,PNStock,COSTA,COSTB,ACTUALAMOUNT,UT2,LASTIN,LASTOUT,"
                SQL = SQL + "DEPONAME,SALES,DISPOSALREASON,ONEYEAR,BUYER,BUYERNAME,CREATEDATE,CREATEUSER)"
                SQL = SQL + " values('" + DAppDate.Text + "','" + DAppName.Text + "','" + DDepoName.Text + "','" + smonth + "','" + sno + "'," + CStr(i + 1) + ","
                For j = 0 To 28


                    If j = 0 Then
                        If GridView1.Rows(i).Cells(j).Text = "&nbsp;" Then
                            jSQL = "''"
                        Else
                            jSQL = "N'" + YKK.ReplaceString(GridView1.Rows(i).Cells(j).Text) + "'"

                        End If

                    Else

                        If GridView1.Rows(i).Cells(j).Text = "&nbsp;" Then
                            jSQL = jSQL + "," + "''"
                        Else
                            jSQL = jSQL + ",N'" + YKK.ReplaceString(GridView1.Rows(i).Cells(j).Text) + "'"

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

            Dim dt_Mail As DataTable
            Dim Adminmail As String = ""
            '���W��
            SQL = "select * from M_referp"
            SQL = SQL + " where cat  ='6001'"
            SQL = SQL + " and dkey = 'adminmail'"
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                Adminmail = dt_Mail.Rows(0).Item("Data").ToString
            End If

            '   Send1(wApplyID, Adminmail, wApplyID, "006001", "1", "1", "END")
            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X, �T�����O

            DInsert.Enabled = False

        Catch ex As Exception
            uJavaScript.PopMsg(Me, "�W���ɮ׮榡���~Insert,�нT�{!")
            inserterror = False
        End Try

        ' End If


        GridView1.DataSource = Nothing
        GridView1.DataBind()
 



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
      
    End Sub


    Protected Sub DUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DUpload.Click

        upload()


    End Sub


 

    Protected Sub DFileUpload1_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs)
        upload()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then

            '�ˬd���W��

            ' s1 = Trim(e.Row.Cells(26).Text.ToUpper)
            DInsert.Enabled = True

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
            ElseIf Trim(e.Row.Cells(27).Text.ToUpper) <> "BUYER" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(28).Text.ToUpper) <> "BUYER NAME" Then
                DInsert.Enabled = False
            End If

        'End If




        If DInsert.Enabled = False Then
            Message = "�W�Ǯ榡���~gridview,�нT�{!"
            uJavaScript.PopMsg(Me, Message)
            griderror = False

            GridView1.DataSource = Nothing
            GridView1.DataBind()

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
            fielderror = False

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


        Dim NewFormSno As Integer = wFormSno    '���y����
        Dim sno As String
        sno = SetNo(NewFormSno)

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
            xStepname = Mid(DAppDate.Text, 1, 6) + "-" + DDepoName.Text + "-" + xFromName + "-" + sno

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

    Protected Sub DITEMNAME1_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME1.TextChanged
        If DITEMNAME1.Text <> "" Then
            DDISRULE1.SelectedIndex = 1
        Else
            DDISRULE1.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME2.TextChanged
        If DITEMNAME2.Text <> "" Then
            DDISRULE2.SelectedIndex = 1
        Else
            DDISRULE2.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME3.TextChanged
        If DITEMNAME3.Text <> "" Then
            DDISRULE3.SelectedIndex = 1
        Else
            DDISRULE3.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME4.TextChanged
        If DITEMNAME4.Text <> "" Then
            DDISRULE4.SelectedIndex = 1
        Else
            DDISRULE4.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME5.TextChanged
        If DITEMNAME5.Text <> "" Then
            DDISRULE5.SelectedIndex = 1
        Else
            DDISRULE5.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME6.TextChanged
        If DITEMNAME6.Text <> "" Then
            DDISRULE6.SelectedIndex = 1
        Else
            DDISRULE6.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME7_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME7.TextChanged
        If DITEMNAME7.Text <> "" Then
            DDISRULE7.SelectedIndex = 1
        Else
            DDISRULE7.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME8_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME8.TextChanged
        If DITEMNAME8.Text <> "" Then
            DDISRULE8.SelectedIndex = 1
        Else
            DDISRULE8.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME9_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME9.TextChanged
        If DITEMNAME9.Text <> "" Then
            DDISRULE9.SelectedIndex = 1
        Else
            DDISRULE9.SelectedIndex = 0
        End If
    End Sub

    Protected Sub DITEMNAME10_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DITEMNAME10.TextChanged
        If DITEMNAME10.Text <> "" Then
            DDISRULE10.SelectedIndex = 1
        Else
            DDISRULE10.SelectedIndex = 0
        End If
    End Sub
 
    Sub DDISRULEChang()
        Dim Rule As String
        Dim i As Integer
        Dim sql As String

        For i = 1 To 10
            Rule = "DDISRULE" + Trim(Str(i))
            Dim DText As DropDownList = Me.FindControl(Rule)
            DText.Items.Clear()
        Next

        If DDISPOSALRULE.SelectedValue = "�C���k��H" Then
            '�Ӷ����o�ǫhappend

            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '6001'"
            sql = sql & " and dkey = 'RULE'"
            sql = sql & " and data = '�C���k'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

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
                    '   DText.SelectedValue = dtr("Data")
                Next

            Next

        Else

            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '6001'"
            sql = sql & " and dkey = 'RULE'"
            sql = sql & " and data = '�D�C���k'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

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
                    ' DText.SelectedValue = dtr("Data")
                Next

            Next

        End If

        If DITEMNAME1.Text <> "" Then
            DDISRULE1.SelectedIndex = 1
        Else
            DDISRULE1.SelectedIndex = 0
        End If


        If DITEMNAME2.Text <> "" Then
            DDISRULE2.SelectedIndex = 1
        Else
            DDISRULE2.SelectedIndex = 0
        End If


        If DITEMNAME3.Text <> "" Then
            DDISRULE3.SelectedIndex = 1
        Else
            DDISRULE3.SelectedIndex = 0
        End If


        If DITEMNAME4.Text <> "" Then
            DDISRULE4.SelectedIndex = 1
        Else
            DDISRULE4.SelectedIndex = 0
        End If


        If DITEMNAME5.Text <> "" Then
            DDISRULE5.SelectedIndex = 1
        Else
            DDISRULE5.SelectedIndex = 0
        End If


        If DITEMNAME6.Text <> "" Then
            DDISRULE6.SelectedIndex = 1
        Else
            DDISRULE6.SelectedIndex = 0
        End If


        If DITEMNAME7.Text <> "" Then
            DDISRULE7.SelectedIndex = 1
        Else
            DDISRULE7.SelectedIndex = 0
        End If


        If DITEMNAME8.Text <> "" Then
            DDISRULE8.SelectedIndex = 1
        Else
            DDISRULE8.SelectedIndex = 0
        End If


        If DITEMNAME9.Text <> "" Then
            DDISRULE9.SelectedIndex = 1
        Else
            DDISRULE9.SelectedIndex = 0
        End If


        If DITEMNAME10.Text <> "" Then
            DDISRULE10.SelectedIndex = 1
        Else
            DDISRULE10.SelectedIndex = 0
        End If

    End Sub

    Protected Sub DSTYPE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSTYPE.SelectedIndexChanged
        If wStep = 1 Then  '�ͺޥD��KEY�߷|���

            Dim sql As String

            If DSTYPE.SelectedValue = "�g������"   Then
                sql = " select convert(char(10),dateadd(day ,-1, dateadd(m, datediff(m,0,getdate())+1,0)),111)lastdate "
            Else
                sql = "SELECT  sdate,edate, convert(char(10),dateadd(day,-1,convert(datetime, sdate, 111)),111)sdate1, "
                sql = sql + " convert(char(10),dateadd(day ,-1, dateadd(m, datediff(m,0,getdate())+1,0)),111)lastdate "
                sql = sql + " FROM ("
                sql = sql + " SELECT  SUBSTRING(DKEY,13,1) AS  DKEYS, DATA AS SDATE   from M_referp"
                sql = sql + " where(cat = 6001)"
                sql = sql + " and left(dkey,12)='SendTime-PCS'"
                sql = sql + " )A, ("
                sql = sql + " SELECT SUBSTRING(DKEY,13,1) AS  DKEYE,DATA AS EDATE   from M_referp"
                sql = sql + " where(cat = 6001)"
                sql = sql + " and left(dkey,12)='SendTime-PCE'"
                sql = sql + " )B "
                sql = sql + " WHERE(A.DKEYS = B.DKEYE)"
                sql = sql + " and left(sdate,7) =left(convert(char(10),getdate(),111),7)  and sdate >= convert(char(10),getdate(),111) "
                sql = sql + " order by sdate "
            End If
           
            Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)

            If dtFlow.Rows.Count > 0 Then

                If DSTYPE.SelectedValue = "�g������" Then
                    DSignDate.Text = dtFlow.Rows(0)("lastdate")
                    ' DSignDate.Text = "�L"
                Else

                    DSignDate.Text = dtFlow.Rows(0)("lastdate")

                End If
            Else
                DSignDate.Text = "���}��"


            End If

            If DSTYPE.SelectedValue = "�릸" Then
                '���o���b�H��
                DPCNAME.Text = GetPCMAN(DAppName.Text)
                If DPCNAME.Text = "" Then
                    uJavaScript.PopMsg(Me, "�L�]�w���b�H���A�лP�ͺ޽T�{")
                End If


            End If

        End If


        

    End Sub

    Function GetPCMAN(ByVal UserName As String) As String

        Dim sql As String

        '= "select  Userid  from M_Users where username=(SELECT SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data)) FROM M_referp where cat ='7002' and Dkey ='RUSER' and Data like '" & UserName & "%') "
        sql = " SELECT * FROM ("
        sql = sql & " SELECT DATA,CASE WHEN SUBSTRING(DATA,1,5) <>'�}����-1' THEN SUBSTRING(DATA,1,CHARINDEX('-',Data)-1)"
        sql = sql & " ELSE SUBSTRING(DATA,1,5) END AS APPNAME"
        sql = sql & " ,CASE WHEN SUBSTRING(DATA,1,5) <>'�}����-1' THEN SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data))"
        sql = sql & " ELSE SUBSTRING(DATA,7,3)END AS   PCMAN"
        sql = sql & " FROM M_referp where cat ='6001' and Dkey ='NEWPCMAN'"
        sql = sql & " )a,m_users b where a.PCMAN=b.username and appname  ='" + UserName + "'"

        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserName")
        End If
        Return NextGate
    End Function

    Function GetUserID(ByVal UserName As String) As String
        Dim sql As String
        sql = " SELECT * FROM M_USERS "
        sql = sql & " WHERE USERNAME ='" + UserName + "'"
      
        Dim dt As Data.DataTable = uDataBase.GetDataTable(Sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("Userid")
        End If
        Return NextGate

    End Function



    Function GetDEPMAN(ByVal UserName As String) As String

        Dim sql As String


        sql = "  SELECT * FROM ( "
        sql = sql & "  SELECT DATA,CASE WHEN SUBSTRING(DATA,1,6 )not in ('������-PL','������-A1','������-DC') THEN SUBSTRING(DATA,1,CHARINDEX('-',Data)-1) "
        sql = sql & " ELSE SUBSTRING(DATA,1,6) END AS APPNAME "
        sql = sql & " ,CASE WHEN SUBSTRING(DATA,1,6) not in ('������-PL','������-A1','������-DC') THEN SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data))"
        sql = sql & " ELSE SUBSTRING(DATA,8,3)END AS   DEPMAN"
        sql = sql & " FROM M_referp where cat ='6001' and Dkey ='DEPMAN'"
        sql = sql & " )a,m_users b where a.DEPMAN=b.username and appname  ='" + UserName + "'"



        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserName")
        End If
        Return NextGate
    End Function


End Class

