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
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID

    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String = "CL"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

    Dim wUserName As String = ""    '姓名代理用
    Dim HolidayList As New List(Of Integer) '用以記錄假日的欄位索引值
    Dim indexList As New List(Of Integer)   '用以記錄不屬於選取月份的欄位索引值
    Dim DateList As New List(Of String)     '用以記錄所選取的一周日期

    ''' <summary>
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式

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
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件


        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

            '上傳資料檢查及顯示訊息

        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號

        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)

        ' wUserID = Request.QueryString("pUserID")


        ' wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))

        Response.Cookies("PGM").Value = "DASW_DISPOSALSheet01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼

       
        ' BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BASDate.Attributes.Add("onClick", "calendarPicker('DASDate')")
        BDISPOSALTYPE.Attributes("onclick") = "GetDISPOSALTYPE();"

        BReason.Attributes.Add("onclick", "window.open('DASW_DISPOSALReasonURL.aspx','','status=0,toolbar=0,width=1000,height=600')")


    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()



        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                             System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")  'WIS-TempPath

        Dim SQL As String
        '主檔資料
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
            LDISPOSALFILE.NavigateUrl = Path & DBAdapter1.Rows(0).Item("DISPOSALFILE1") '報廢明細
            LDISPOSALFILE.Visible = True
        Else
            LDISPOSALFILE.Visible = False
        End If

        DPCNAME.Text = DBAdapter1.Rows(0).Item("PCNAME")

        'SetFieldData("PCNAME", DBAdapter1.Rows(0).Item("PCNAME"))


        If Mid(DBAdapter1.Rows(0).Item("SignDate").ToString, 1, 4) = "1900" Then
            DSignDate.Text = ""
        Else
            DSignDate.Text = DBAdapter1.Rows(0).Item("SignDate1")             '簽核時間
        End If

        '立會時間

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

     
        If wStep = 500 Then  '生管主管KEY立會日期

            If DSignDate.Text <> "無" Then

                If DSignDate.Text <> "未開放" Then
                    Dim SQL1 As String
                    '主檔資料
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
                DASDate.Value = DBAdapter1.Rows(0).Item("ASDate")               '立會時間
            End If
        End If




        'Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '6001'"
        SQL = SQL + " and dkey ='DisposalFilePath'"
        Dim DBAdapter6 As DataTable = uDataBase.GetDataTable(SQL)
        Dim OpenDir1 As String = ""
        If DBAdapter6.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter6.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter6.Rows(0).Item("Data1")
        End If


        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If


        '開啟報廢資料夾路徑
        DDISPOSALFILE2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "\" + DNo.Text + "','_blank');return true;")




        '明細資料
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



        GetTotal() '重新計算


        Dim SIGN, SIGNDATE As String
        '寫入簽核主管資料
        SQL = " select rank() over(order by [step]) as cno,decidename,convert(char(10),completedtime,111)CompletedDate "
        SQL = SQL & " from (select  step,decidename,max(completedtime)completedtime  from  t_waithandle"
        SQL = SQL & " where  FORMNO = '006001' AND  FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & " And Step <= 100  And Step not in (1,500,999)  and sts =1 and flowtype <> 0"
        SQL = SQL & " group by  step,decidename"
        SQL = SQL & " )a order by step "
        Dim DBAdapter4 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter4.Rows.Count > 0 Then

            For i = 0 To DBAdapter4.Rows.Count - 1
                '簽核主管
                SIGN = "DSIGN" + Trim(Str(i + 1))
                Dim DSIGN As TextBox = Me.FindControl(SIGN)
                DSIGN.Text = DBAdapter4.Rows(i).Item("decidename")
                '簽核日期
                SIGNDATE = "DSIGNDATE" + Trim(Str(i + 1))
                Dim DSIGNDATE As TextBox = Me.FindControl(SIGNDATE)
                DSIGNDATE.Text = DBAdapter4.Rows(i).Item("CompletedDate")

            Next
        End If


        getDUTYDEPOID()





        '寫入流程資料
        SQL = "Select * From T_WaitHandle "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
        SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"

        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter3.Rows.Count > 0 Then
            DDecideDesc.Text = DBAdapter3.Rows(0).Item("DecideDesc")       '說明


            If DBAdapter3.Rows(0).Item("BEndTime") < NowDateTime Then
                If DDelay.Visible = True Then
                    SetFieldData("ReasonCode", DBAdapter3.Rows(0).Item("ReasonCode"))    '延遲理由代碼
                    If DBAdapter3.Rows(0).Item("ReasonCode") = "" Then
                        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                        DReasonDesc.Text = ""   '延遲其他說明
                    Else
                        DReason.Text = DBAdapter3.Rows(0).Item("Reason")  '延遲理由
                        DReasonDesc.Text = DBAdapter3.Rows(0).Item("ReasonDesc")     '延遲其他說明
                    End If
                End If
            End If

        End If






        '核定履歷資料
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

        'DB連結關閉
        'OleDbConnection1.Close()

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String

        'DB連結設定
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)


        If DBAdapter1.Rows.Count > 0 Then
            '電子簽章未使用
            If DBAdapter1.Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBAdapter1.Rows(0).Item("Attach") = 1 Then
            Else
            End If
            '儲存按鈕
            If DBAdapter1.Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBAdapter1.Rows(0).Item("SaveDesc")
                BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBAdapter1.Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBAdapter1.Rows(0).Item("NGDesc1")
                BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                wNGSts1 = DBAdapter1.Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBAdapter1.Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBAdapter1.Rows(0).Item("NGDesc2")
                BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                wNGSts2 = DBAdapter1.Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBAdapter1.Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBAdapter1.Rows(0).Item("OKDesc")
                BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                wOKSts = DBAdapter1.Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If DBAdapter1.Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then

                '遲納管理
                If wDelay = 1 Then
                    If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                    Else
                        DDelay.Visible = False  '延遲Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '延遲理由代碼
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                    DReason.Visible = True         '延遲理由
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                    DReasonDesc.Visible = True     '延遲其他說明
                    Top = 1220
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 1220
                End If



                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If DBAdapter2.Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If

                '按鈕位置
                BNG1.Style.Add("Top", Top)      'NG按鈕
                BNG2.Style.Add("Top", Top)     'NG按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕

            End If
        Else
            Top = 1150
            'Sheet隱藏

            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False     '說明
            DDescSheet.Visible = False
            DReasonCode.Visible = False     '延遲理由代碼
            DReason.Visible = False         '延遲理由
            DReasonDesc.Visible = False     '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)      'NG按鈕
            BNG2.Style.Add("Top", Top)     'NG按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕

        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        '日期選擇



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     延遲理由代碼點選後事件
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String


        '按鈕及RequestedField的Top位置
        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
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
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性

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
        Select Case FindFieldInf("NO")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
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
        Select Case FindFieldInf("AppDate")
            Case 0  '顯示
                DAppDate.BackColor = Color.LightGray
                DAppDate.ReadOnly = True
                DAppDate.Visible = True

            Case 1  '修改+檢查
                DAppDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DAppDate.Visible = True

            Case 2  '修改
                DAppDate.BackColor = Color.Yellow
                DAppDate.Visible = True

            Case Else   '隱藏
                DAppDate.Visible = False

        End Select
        If pPost = "New" Then DAppDate.Text = Now.ToString("yyyy/MM/dd") '現在日時

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '修改+檢查
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入部門")
                DDepoName.Visible = True
            Case 2  '修改
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DepoName", "ZZZZZZ")


        'Name
        Select Case FindFieldInf("AppName")
            Case 0  '顯示
                DAppName.BackColor = Color.LightGray
                DAppName.ReadOnly = True
                DAppName.Visible = True
            Case 1  '修改+檢查
                DAppName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAPPNameRqd", "DAPPName", "異常：需輸入姓名")
                DAppName.Visible = True
            Case 2  '修改
                DAppName.BackColor = Color.Yellow
                DAppName.Visible = True
            Case Else   '隱藏
                DAppName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AppName", "ZZZZZZ")



        '申請原因
        Select Case FindFieldInf("DISPOSALREASON")
            Case 0  '顯示
                DDISPOSALREASON.BackColor = Color.LightGray
                DDISPOSALREASON.Visible = True

            Case 1  '修改+檢查
                DDISPOSALREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDISPOSALREASONRqd", "DDISPOSALREASON", "異常：需輸申請原因")
                DDISPOSALREASON.Visible = True
            Case 2  '修改
                DDISPOSALREASON.BackColor = Color.Yellow
                DDISPOSALREASON.Visible = True
            Case Else   '隱藏
                DDISPOSALREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DISPOSALREASON", "ZZZZZZ")


        '其它原因
        Select Case FindFieldInf("DISPOSALREASON1")
            Case 0  '顯示
                DDISPOSALREASON1.BackColor = Color.LightGray
                DDISPOSALREASON1.ReadOnly = True
                DDISPOSALREASON1.Visible = True
            Case 1  '修改+檢查
                DDISPOSALREASON1.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DDISPOSALREASON1Rqd", "DDISPOSALREASON1", "異常：需輸入其它原因")
                DDISPOSALREASON1.Visible = True
            Case 2  '修改
                DDISPOSALREASON1.BackColor = Color.Yellow
                DDISPOSALREASON1.Visible = True
            Case Else   '隱藏
                DDISPOSALREASON1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DISPOSALREASON1", "ZZZZZZ")




        '業務
        Select Case FindFieldInf("Sales")
            Case 0  '顯示
                DSales.BackColor = Color.LightGray
                DSales.ReadOnly = True
                DSales.Visible = True
            Case 1  '修改+檢查
                DSales.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DSalesRqd", "DSales", "異常：需輸入其它原因")
                DSales.Visible = True
            Case 2  '修改
                DSales.BackColor = Color.Yellow
                DSales.Visible = True
            Case Else   '隱藏
                DSales.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sales", "ZZZZZZ")


        '報廢形式
        Select Case FindFieldInf("STYPE")
            Case 0  '顯示
                DSTYPE.BackColor = Color.LightGray
                DSTYPE.Visible = True

            Case 1  '修改+檢查
                DSTYPE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSTYPERqd", "DSTYPE", "異常：需輸入報廢形式")
                DSTYPE.Visible = True
            Case 2  '修改
                DSTYPE.BackColor = Color.Yellow
                DSTYPE.Visible = True
            Case Else   '隱藏
                DSTYPE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("STYPE", "ZZZZZZ")



        '責任部門
        Select Case FindFieldInf("DUTYDEPO")
            Case 0  '顯示
                DDUTYDEPO.BackColor = Color.LightGray
                DDUTYDEPO.Visible = True

            Case 1  '修改+檢查
                DDUTYDEPO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDUTYDEPORqd", "DDUTYDEPO", "異常：需輸入責任部門")
                DDUTYDEPO.Visible = True
            Case 2  '修改
                DDUTYDEPO.BackColor = Color.Yellow
                DDUTYDEPO.Visible = True
            Case Else   '隱藏
                DDUTYDEPO.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DUTYDEPO", "ZZZZZZ")

        '場所
        Select Case FindFieldInf("PLACE")
            Case 0  '顯示
                DPLACE.BackColor = Color.LightGray
                DPLACE.Visible = True

            Case 1  '修改+檢查
                DPLACE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPLACERqd", "DPLACE", "異常：需輸入資產在庫場所")
                DPLACE.Visible = True
            Case 2  '修改
                DPLACE.BackColor = Color.Yellow
                DPLACE.Visible = True
            Case Else   '隱藏
                DPLACE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PLACE", "ZZZZZZ")

        '報廢準則
        Select Case FindFieldInf("DISPOSALRULE")
            Case 0  '顯示
                DDISPOSALRULE.BackColor = Color.LightGray
                DDISPOSALRULE.Visible = True

            Case 1  '修改+檢查
                DDISPOSALRULE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDISPOSALRULERqd", "DDISPOSALRULE", "異常：需輸入報廢準則")
                DDISPOSALRULE.Visible = True
            Case 2  '修改
                DDISPOSALRULE.BackColor = Color.Yellow
                DDISPOSALRULE.Visible = True
            Case Else   '隱藏
                DDISPOSALRULE.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DISPOSALRULE", "ZZZZZZ")



        'Name
        Select Case FindFieldInf("AppName")
            Case 0  '顯示
                DAppName.BackColor = Color.LightGray
                DAppName.ReadOnly = True
                DAppName.Visible = True
            Case 1  '修改+檢查
                DAppName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAPPNameRqd", "DAPPName", "異常：需輸入姓名")
                DAppName.Visible = True
            Case 2  '修改
                DAppName.BackColor = Color.Yellow
                DAppName.Visible = True
            Case Else   '隱藏
                DAppName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AppName", "ZZZZZZ")




        'DisposalType
        Select Case FindFieldInf("DISPOSALTYPE")
            Case 0  '顯示
                DDISPOSALTYPE.Visible = True
                DDISPOSALTYPE.Style.Add("background-color", "lightgrey")
                DDISPOSALTYPE.Attributes.Add("readonly", "true")
                BDISPOSALTYPE.Visible = False

            Case 1  '修改+檢查
                DDISPOSALTYPE.Visible = True
                DDISPOSALTYPE.Style.Add("background-color", "greenyellow")
                DDISPOSALTYPE.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DDISPOSALTYPERqd", "DDISPOSALTYPE", "異常：需輸入報廢品類別")
            Case 2  '修改
                DDISPOSALTYPE.Visible = True
                DDISPOSALTYPE.Style.Add("background-color", "yellow")
                DDISPOSALTYPE.Attributes.Add("readonly", "true")
                BDISPOSALTYPE.Visible = True
            Case Else   '隱藏
                DDISPOSALTYPE.Visible = False
                BDISPOSALTYPE.Visible = False

        End Select
        If pPost = "New" Then DDISPOSALTYPE.Value = ""





        'CHINESEREASON
        Select Case FindFieldInf("CHINESEREASON")
            Case 0  '顯示
                DCHINESEREASON.BackColor = Color.LightGray
                DCHINESEREASON.ReadOnly = True
                DCHINESEREASON.Visible = True
            Case 1  '修改+檢查
                DCHINESEREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCHINESEREASONRqd", "DCHINESEREASON", "異常：需輸入中文原因")
                DCHINESEREASON.Visible = True
            Case 2  '修改
                DCHINESEREASON.BackColor = Color.Yellow
                DCHINESEREASON.Visible = True
            Case Else   '隱藏
                DCHINESEREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CHINESEREASON", "ZZZZZZ")


        'JAPANREASON
        Select Case FindFieldInf("JAPANREASON")
            Case 0  '顯示
                DJAPANREASON.BackColor = Color.LightGray
                DJAPANREASON.ReadOnly = True
                DJAPANREASON.Visible = True
            Case 1  '修改+檢查
                DJAPANREASON.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJAPANREASONRqd", "DJAPANREASON", "異常：需輸入日文原因")
                DJAPANREASON.Visible = True
            Case 2  '修改
                DJAPANREASON.BackColor = Color.Yellow
                DJAPANREASON.Visible = True
            Case Else   '隱藏
                DJAPANREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("JAPANREASON", "ZZZZZZ")


        Dim i As Integer
        Dim ITEMCODE As String
        Select Case FindFieldInf("ITEMCODE")
            Case 0  '顯示
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DITEMCODE" + Str(i) + "Rqd", ITEMCODE, "異常：需輸入ITEMCODE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 10
                    ITEMCODE = "DITEMCODE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMCODE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DITEMNAME" + Str(i) + "Rqd", ITEMNAME, "異常：需輸入ITEMNAME" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 10
                    ITEMNAME = "DITEMNAME" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(ITEMNAME)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DPIECE" + Str(i) + "Rqd", PIECE, "異常：需輸入PIECE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 13
                    PIECE = "DPIECE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PIECE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DMETER" + Str(i) + "Rqd", METER, "異常：需輸入METER" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 13
                    METER = "DMETER" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(METER)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DYARD" + Str(i) + "Rqd", YARD, "異常：需輸入YARD" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 13
                    YARD = "DYARD" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(YARD)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DKG" + Str(i) + "Rqd", KG, "異常：需輸入KG" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 13
                    KG = "DKG" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(KG)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DPRICE" + Str(i) + "Rqd", PRICE, "異常：需輸入PRICE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 13
                    PRICE = "DPRICE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(PRICE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                DSUM.Visible = True
            Case 1  '修改+檢查

                DSUM.Visible = True
            Case 2  '修改

                DSUM.Visible = True
            Case Else   '隱藏
                DSUM.Visible = False
        End Select



        Dim DDISRULE As String
        Select Case FindFieldInf("RULE")
            Case 0  '顯示
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.BackColor = Color.LightGray
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDISRULE " + Str(i) + "Rqd", DDISRULE, "異常：需輸入報廢規則" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 10
                    DDISRULE = "DDISRULE" + Trim(Str(i))
                    Dim DText As DropDownList = Me.FindControl(DDISRULE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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
            Case 0  '顯示
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DREMARK" + Str(i) + "Rqd", REMARK, "異常：需輸入REMARK" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 10
                    REMARK = "DREMARK" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(REMARK)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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




        '生管扣帳
        Select Case FindFieldInf("PCNAME")
            Case 0  '顯示
                DPCNAME.BackColor = Color.LightGray
                DPCNAME.Visible = True
                DPCNAME.ReadOnly = True
            Case 1  '修改+檢查
                DPCNAME.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPCNAMERqd", "DPCNAME", "異常：需輸入生管扣帳擔當人員")
                DPCNAME.Visible = True
            Case 2  '修改
                DPCNAME.BackColor = Color.Yellow
                DPCNAME.Visible = True
            Case Else   '隱藏
                DPCNAME.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PCNAME", "ZZZZZZ")


        '報廢明細
        Select Case FindFieldInf("DISPOSALFILE1")
            Case 0  '顯示
                DFileUpload1.Visible = False
                DFileUpload1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ' ShowRequiredFieldValidator("DFileUpload1Rqd", "DFileUpload1", "異常：需上傳報廢附檔")
                DFileUpload1.Visible = True
                DFileUpload1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DFileUpload1.Visible = True
                DFileUpload1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DFileUpload1.Visible = False
        End Select


        '報廢附檔
        Select Case FindFieldInf("DISPOSALFILE2")
            Case 0  '顯示
                DDISPOSALFILE2.Enabled = False
            Case 1  '修改+檢查

                DDISPOSALFILE2.Visible = True

            Case 2  '修改
                DDISPOSALFILE2.Visible = True

            Case Else   '隱藏
                DDISPOSALFILE2.Visible = False
        End Select



        '簽核主管
        Dim SIGN As String
        Select Case FindFieldInf("SIGN")
            Case 0  '顯示
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DSIGN" + Str(i) + "Rqd", SIGN, "異常：需輸入SIGN" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 12
                    SIGN = "DSIGN" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGN)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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


        '簽核主管
        Dim SIGNDATE As String
        Select Case FindFieldInf("SIGNDATE")
            Case 0  '顯示
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.BackColor = Color.LightGray
                    DText.ReadOnly = True
                    DText.Visible = True
                Next

            Case 1  '修改+檢查
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DSIGNDATE" + Str(i) + "Rqd", SIGNDATE, "異常：需輸入SIGNDATE" + Str(i))
                    DText.Visible = True
                Next

            Case 2  '修改
                For i = 1 To 12
                    SIGNDATE = "DSIGNDATE" + Trim(Str(i))
                    Dim DText As TextBox = Me.FindControl(SIGNDATE)
                    DText.BackColor = Color.Yellow
                    DText.Visible = True
                Next

            Case Else   '隱藏
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




        '營業簽核
        Select Case FindFieldInf("MKTSIGN")
            Case 0  '顯示
                ' DMKTSIGN.BackColor = Color.LightGray
                DMKTSIGN.Enabled = False

            Case 1  '修改+檢查
                DMKTSIGN.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMKTSIGNRqd", "DMKTSIGN", "異常：需輸入需營業簽核")

            Case 2  '修改
                DMKTSIGN.BackColor = Color.LightGreen

            Case Else   '隱藏
                DMKTSIGN.Visible = False
        End Select
        If pPost = "New" Then DMKTSIGN.Checked = False


        '向客戶取款
        Select Case FindFieldInf("CUSTOMERTOLL")
            Case 0  '顯示
                ' DCUSTOMERTOLL.BackColor = Color.LightGray
                DCUSTOMERTOLL.Enabled = False

            Case 1  '修改+檢查
                DCUSTOMERTOLL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCUSTOMERTOLLRqd", "DCUSTOMERTOLL", "異常：需輸入向客戶取款")

            Case 2  '修改
                DCUSTOMERTOLL.BackColor = Color.LightGreen

            Case Else   '隱藏
                DCUSTOMERTOLL.Visible = False
        End Select
        If pPost = "New" Then DCUSTOMERTOLL.Checked = False



        '立會日期
        'DeliveryDate
        Select Case FindFieldInf("ASDate")
            Case 0  '顯示
                DASDate.Visible = True
                DASDate.Style.Add("background-color", "lightgrey")
                DASDate.Attributes.Add("readonly", "true")
                BASDate.Visible = False

            Case 1  '修改+檢查
                DASDate.Visible = True
                DASDate.Style.Add("background-color", "greenyellow")
                DASDate.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DASDateRqd", "DASDate", "異常：需輸入立會日期")
                BASDate.Visible = True

            Case 2  '修改
                DASDate.Visible = True
                DASDate.Style.Add("background-color", "yellow")
                DASDate.Attributes.Add("readonly", "true")
                BASDate.Visible = True
            Case Else   '隱藏
                DASDate.Visible = False
                BASDate.Visible = False

        End Select
        If pPost = "New" Then DASDate.Value = ""



        'If pPost = "New" Then DASDate.Value = Now.ToString("yyyy/MM/dd") '現在日時

        '立會日期
        Select Case FindFieldInf("SignDate")
            Case 0  '顯示
                DSignDate.BackColor = Color.LightGray
                DSignDate.ReadOnly = True
                DSignDate.Visible = True

            Case 1  '修改+檢查
                DSignDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSignDateRqd", "DSignDate", "異常：需輸入簽核日期")
                DSignDate.Visible = True

            Case 2  '修改
                DSignDate.BackColor = Color.Yellow
                DSignDate.Visible = True

            Case Else   '隱藏
                DSignDate.Visible = False

        End Select
        'If pPost = "New" Then DSignDate.Text = Now.ToString("yyyy/MM/dd") '現在日時


        '是否需立會
        Select Case FindFieldInf("CheckASDate")
            Case 0  '顯示
                ' DCUSTOMERTOLL.BackColor = Color.LightGray
                CheckASDate.Enabled = False

            Case 1  '修改+檢查
                CheckASDate.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("CblRqd", "Cbl", "異常：需輸入是否需立會")

            Case 2  '修改
                CheckASDate.BackColor = Color.Yellow

            Case Else   '隱藏
                CheckASDate.Visible = False
        End Select
        'If pPost = "New" Then cbl_1.Checked = False

        If wStep = 1 Then  '生管主管KEY立會日期
            If DSignDate.Text <> "無" Then
                If DSignDate.Text <> "未開放" Then
                    Dim SQL1 As String
                    '主檔資料
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
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.None              ' 因在頁面上加入ValidationSummary , 故驗證控制項統一顯示
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)



    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer



        idx = FindFieldInf(pFieldName)

        '擔當者及部門 

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


        '取得扣帳人員
        DPCNAME.Text = GetPCMAN(DAppName.Text)


        '簽核期限
        If DSTYPE.SelectedValue = "月次" Or wStep = 1 Then

            If checkDate() = 1 Then
                DSignDate.Text = SignDate

            Else
                If wStep = 1 Then
                    DSignDate.Text = "未開放"
                End If

            End If

            If DSTYPE.SelectedValue = "週次尾數" Then
                DSignDate.Text = "無"
            End If

        End If




        '申請原因
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

        '責任部門
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

        '在庫場所
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


        '報廢型式
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



        '報廢準則
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


        '細項報廢準則append
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

                '新增RULE
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

        '細項報廢準則showfomrdata  


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

            '新增RULE
            For Each dtr As Data.DataRow In DBAdapter1.Rows


                Rule = "DDISRULE" + Trim(Str(i))
                DText.Items.Add(dtr("Data"))

            Next


        End If



        ''生管扣帳人員
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
        '    '生管扣帳人員
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
    '**     尋找表單欄位屬性
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
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True And OK() Then


            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")


        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG1按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")


        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG2按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then

            If wStep = 10 Then
                If OK() = True Then
                    DisabledButton()   '停止Button運作
                    FlowControl("NG2", 2, "3")

                End If
            Else
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")

            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T


        GetDataStatus()  '取得此關卡各按鈕的Data Status



        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("006001", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '樣品難易度

            While Run = True
                Run = False     '執行Flag=不執行
                '--取得下一關參數設定---------
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=通知
                Dim pCount As Integer
                '--取得LeadTime參數設定---------
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--取得工程負荷參數設定---------
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--其他參數設定---------
                Dim i, RtnCode As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
                Dim SQL As String = ""

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        '取得表單流水號
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else

                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo,Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End
                            '設定委託No
                            DNo.Text = SetNo(NewFormSno)

                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            ModifyTranData(pFun, pSts)

                            '流程資料結束
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                        End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then

                    Dim wAllocateID As String = ""

                    If (wStep = 1 Or wStep = 500) And pFun = "OK" And DSTYPE.SelectedValue = "月次" Then
                        wAllocateID = GetRelated(wApplyID)  '台藉
                        pAction = 0
                        If DDISPOSALTYPE.Value <> "DEPOT-07" Then
                            pAction = 0
                        Else
                            pAction = 2
                        End If

                    End If

                    If (wStep = 1 Or wStep = 502) And pFun = "OK" And DSTYPE.SelectedValue = "週次尾數" Then
                        pAction = 1
                    End If

                    If wStep = 6 Then
                        wAllocateID = GetRelated(wApplyID)  '台藉
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


                    '找出簽核者User ID 台藉主管 
                    If wStep = 111 Then  '指定台藉主管 
                        If pAction = 0 Then
                            wAllocateID = GetRelated(wApplyID)
                        End If
                    End If


                    '找出簽核者User ID  
                    If wStep = 112 Then  '指定生管擔當扣帳
                        If pAction = 0 Then
                            wAllocateID = GetUserID(DPCNAME.Text)
                        End If
                    End If





                    '20240211 取消此規定
                    ''如果是生管的人起單就不要再跑生管簽核
                    'If DDepoName.Text = "生產管理" And wStep = 20 And pFun = "OK" Then
                    '    If DMKTSIGN.Checked = True Then '營業
                    '        pAction = 2
                    '    Else
                    '        pAction = 3  '責任部門
                    '        wAllocateID = DDUTYDEPOID.Text
                    '    End If
                    'End If

                    If wStep = 20 And pAction = 1 Then
                        wAllocateID = GetRelated(wApplyID)
                    End If

                    If wStep = 45 Then

                        If DMKTSIGN.Checked = True Then
                            '取得簽核者 ID  
                            wAllocateID = getDUTYDEPOID(wFormSno)

                            '增加CC 絡經理
                            If DDISPOSALTYPE.Value = "PN" Or DDISPOSALTYPE.Value = "QF" Or DDISPOSALTYPE.Value = "S&B" Then
                                pAction = 3
                            Else
                                If DDUTYDEPO.SelectedValue = "營業五課" Then
                                    pAction = 0
                                Else
                                    pAction = 2
                                End If

                                '如果勾營業簽核要跳到50責任部門
                            End If

                        Else
                            '取得簽核者 ID  
                            '如果勾營業簽核要跳到50責任部門
                            wAllocateID = getDUTYDEPOID(wFormSno)


                            '增加CC 絡經理
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
                            wAllocateID = GetRelated(GetRelated(wApplyID)) '提出部門日籍主管

                            If wAllocateID.ToUpper = "PC028" Then
                                pAction = 3
                            Else
                                pAction = 3
                            End If
                        End If

                    End If


                    '如果沒勾營業簽核要跳到60責任部門
                    If wStep = 50 And pFun = "NG1" Then
                        '取得簽核者 ID

                        If DDepoName.Text = "生產管理" Then
                            pAction = 2
                            wAllocateID = GetRelated(wApplyID)
                        Else
                            pAction = 1

                        End If

                    End If

                    '20151223 增加CC營業蔡經理 
                    If wStep = 51 And pFun = "OK" Then

                        If DMKTSIGN.Checked = True Then
                            wAllocateID = GetRelated(getDUTYDEPOID(wFormSno))
                            If wAllocateID = "mk002" Then
                                If DDUTYDEPO.SelectedValue <> "營業五課" Then
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
                    ' wAllocateID = GetRelated(GetRelated(wApplyID)) '提出部門日籍主管
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
                        wAllocateID = GetRelated(GetRelated(wApplyID)) '責任部門主管1
                    End If


                    If wStep = 80 And pFun = "NG1" Then
                        wAllocateID = GetRelated(GetRelated(GetRelated(wApplyID))) '責任部門主管2
                        pAction = 1
                    End If



                    '如果有勾營業簽核又NG要先跳回營業
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
                        wAllocateID = GetRelated(GetRelated(wApplyID)) '提出部門日籍主管
                    End If


             

                    '找出實際處理人員 ID  
                    If wStep = 115 Then   '取得實際處理人員
                        If pAction = 0 Then
                            wAllocateID = GetUserID(GetDEPMAN(DAppName.Text))
                        End If
                    End If




                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                    If ErrCode = 0 Then pAction = 0
                End If

                '建置流程資料
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        wNextGate = ""
                        For i = 1 To pCount
                            '取得下一關人員(訊息時使用)
                            If wNextGate = "" Then
                                wNextGate = pNextGate(i)
                            Else
                                wNextGate = "," & pNextGate(i)
                            End If

                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo,Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '取得核定者-群組行是曆
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '取得工程負荷最後日時
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數

                            '取得預定開始,完成日程計算
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆

                            '建置流程資料
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo,Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        'Modify-End
                    End If
                End If
                '儲存表單資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If  'pSeqno <> 0
                    Else    '判斷是否起單
                        If pNextStep = 999 Then     '工程結束嗎?


                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                        Else

                            ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            If pNextStep <> 30 And pNextStep <> 140 And pNextStep <> 130 And pNextStep <> 80 And wStep <> 70 And wStep <> 80 Then '廠長 竹本 總經理不收信
                                oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                            End If
                            '    oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                        Next i
                    Else
                        oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
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



                    wStep = pNextStep     '下一工程關卡號碼

                    wFormSno = NewFormSno '下一工程表單流水號
                Else
                    EnabledButton()   '起動Button運作
                    If ErrCode = 9110 Then Message = "取得表單流水號計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9120 Then Message = "流程資料更新異常,請確認或連絡系統人員!"
                    If ErrCode = 9130 Then Message = "下工程計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9131 Then Message = "無下工程管理人,請確認或連絡系統人員!"
                    If ErrCode = 9140 Then Message = "工程預定開始及完成日計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9150 Then Message = "下一工程資料建置異常,請確認或連絡系統人員!"
                    If ErrCode = 9160 Then Message = "工程結束資料建置異常(999),請確認或連絡系統人員!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '儲存表單ErrCode=0
            End While       '重覆執行


            If ErrCode = 0 Then
                '--郵件傳送---------
                oCommon.SendMail()

                Dim URL As String



                URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                           "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID

                Response.Redirect(URL)


            End If



        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "材質為其他時需填寫說明,請確認!"
            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9070 Then Message = "需填入工程待處理部門及工程待處理者.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     取得表單進度狀態
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
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
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
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '006001', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號
        If DSignDate.Text = "無" Then
            SQl = SQl + " '" + Mid(Now.ToString("yyyy/MM/dd"), 1, 7) + "', "   '報廢年月
        Else

            SQl = SQl + " '" + Mid(DSignDate.Text, 1, 7) + "', "   '報廢年月
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

        SQl = SQl + " N'" + DAppName.Text + "', "                '申請者 
        SQl = SQl + " N'" + DDepoName.Text + "', "                '部門
        SQl = SQl + " N'" + DAppDate.Text + "', "                '日期
        SQl = SQl + " N'" + DDISPOSALREASON.SelectedValue + "', "                '報廢原因
        SQl = SQl + " N'" + DDISPOSALREASON1.Text + "', "                '報廢原因
        SQl = SQl + " N'" + DDUTYDEPO.SelectedValue + "', "                '責任部門
        SQl = SQl + " N'" + DSales.Text + "', "                '業務
        SQl = SQl + " N'" + DPLACE.SelectedValue + "', "                '場所
        SQl = SQl + " N'" + DSTYPE.SelectedValue + "', "                '場所
        SQl = SQl + " N'" + DDISPOSALRULE.SelectedValue + "', "                '準則
        SQl = SQl + " N'" + DDISPOSALTYPE.Value + "', "                '分類
        SQl = SQl + " N'" + YKK.ReplaceString(DCHINESEREASON.Text) + "', "                '報廢中文
        SQl = SQl + " N'" + YKK.ReplaceString(DJAPANREASON.Text) + "', " '報廢日文
        '姓名4
        FileName = ""
        If DFileUpload1.PostedFile.FileName <> "" Then
            Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                          CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
            Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                        System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")

            '20170912 將檔名修改成不含原始檔名
            Dim fileExtension As String  '副檔名
            fileExtension = IO.Path.GetExtension(DFileUpload1.PostedFile.FileName).ToLower   '取得檔案格式

            FileName = CStr(NewFormSno) + "-DPFILE-" + UploadDateTime + fileExtension

            DFileUpload1.PostedFile.SaveAs(Path + FileName)

        Else
            FileName = ""
        End If

        SQl = SQl + "'" + FileName + "'," '報廢明細
        SQl = SQl + " N'" + DPCNAME.Text + "', "                '生管扣帳
        SQl = SQl + " '" + DASDate.Value + "', "                '立會時間
        '立會時間
        Dim j As Integer
        For j = 0 To CheckASDate.Items.Count - 1 Step j + 1
            ' check item selected state.
            If (CheckASDate.Items(j).Selected) Then
                SQl = SQl + CheckASDate.Items(j).Value + ","
                Exit For
            End If

        Next

        If DSTYPE.SelectedValue = "月次" And DDISPOSALTYPE.Value <> "DEPOT-07" Then
            SQl = SQl + " '" + DSignDate.Text + "', "                '簽核期限
        Else '週次尾數預設每月最後一天
            SQl = SQl + "CONVERT(CHAR(10),dateadd(day,-day(getdate( )),dateadd(month,1,getdate())),111), "                '簽核期限
        End If





        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "        '作成者
        SQl = SQl + " '" + NowDateTime + "', "                            '作成時間
        SQl = SQl + " '" + "" + "', "                                     '修改者
        SQl = SQl + " '" + NowDateTime + "' "                             '修改時間
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)


        '存入受支援明細

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
                sql2 = sql2 + " '" + Request.QueryString("pUserID") + "', "        '作成者
                sql2 = sql2 + " '" + NowDateTime + "', "                            '作成時間
                sql2 = sql2 + " '" + "" + "', "                                     '修改者
                sql2 = sql2 + " '" + NowDateTime + "' "                             '修改時間
                sql2 = sql2 + " ) "
                uDataBase.ExecuteNonQuery(sql2)
            End If


        Next




    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)

        Dim SQL As String = ""

        '主檔資料
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
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                            System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")

                '20170912 將檔名修改成不含原始檔名
                Dim fileExtension As String  '副檔名
                fileExtension = IO.Path.GetExtension(DFileUpload1.PostedFile.FileName).ToLower   '取得檔案格式

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
            SQL = SQL + " DISPOSALFILE1='" + FileName + "'," '報廢證明
        End If
        SQL = SQL + " PIECE ='" + DPIECE13.Text + "',"
        SQL = SQL + " METER ='" + DMETER13.Text + "',"
        SQL = SQL + " YARD ='" + DYARD13.Text + "',"
        SQL = SQL + " KG ='" + DKG13.Text + "',"
        SQL = SQL + " PRICE ='" + DPRICE13.Text + "',"
        SQL = SQL + " PCNAME= N'" & DPCNAME.Text & "',"
        SQL = SQL + " ASDate= '" & DASDate.Value & "',"
        SQL = SQL + " SignDate= '" & DSignDate.Text & "',"



        '立會時間
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

        '明細資料



        '存入受支援明細


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
                sql2 = sql2 + " '" + Request.QueryString("pUserID") + "', "        '作成者
                sql2 = sql2 + " '" + NowDateTime + "', "                            '作成時間
                sql2 = sql2 + " '" + "" + "', "                                     '修改者
                sql2 = sql2 + " '" + NowDateTime + "' "                             '修改時間
                sql2 = sql2 + " ) "
                uDataBase.ExecuteNonQuery(sql2)
            End If


        Next



    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
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
    '**     追加交易資料和委託單對照表
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
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '定義允許的檔案格式
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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
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
    '**     起動Button(OK, Ng1, NG2, Save)運作 
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
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()



        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top - 140 & "px"
            DDecideDesc.Style("top") = Top - 140 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
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
        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編製委託No 
    '**
    '*****************************************************************
    Function SetNo(ByVal Seq As Integer) As String
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim i As Integer

        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '日
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = "D" + CStr(DateTime.Now.Year) + Str2
        '年
        'Set單號
        '當月份筆數有幾筆  20150414 Modify by Jessica
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
    '**     輸入檢查
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


        '判斷是否為可申請報廢

        If wStep = 1 Then
            If DSTYPE.Text = "月次" Then
                If checkDate() = 0 Then
                    isOK = False
                    Message = "異常：" + "已過申請期限:" + SignDate + ",下次開放申請日期:" + OpenDate

                End If
            End If

        End If

        '20241231

        If DSignDate.Text = "未開放" Then
            Message = "異常：簽核期限=未開放不能申請!"
            isOK = False
        End If




        '如果選其它要說明理由



        Dim I As Integer
        Dim ITEMNAME, PIECE, METER, YARD, KG, PRICE, DISRULE, wRow As String

        For I = 1 To 10

            wRow = "第" + CStr(I) + "列"

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


            '檢查 數量否有資料
            If I = 1 Then
                If DITEMNAME.Text = "" Then
                    isOK = False
                    Message = "異常：需輸入ITEMNAME!"
                    Exit For
                End If

            End If


            '檢查 ITEMNAME 是否有資料
            If DDISRULE.SelectedValue <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DKG.Text <> "" Or DPRICE.Text <> "" Or DYARD.Text <> "" Then
                If DITEMNAME.Text = "" Then
                    isOK = False
                    Message = "異常：需輸入ITEMNAME!"
                End If
            End If

            If DDISRULE.SelectedValue <> "" Or DITEMNAME.Text <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DPRICE.Text <> "" Or DYARD.Text <> "" Then
                If DKG.Text = "" Then '檢查 重量否有資料
                    isOK = False
                    Message = Message + "\n" + "異常：需輸入重量!"
                End If
            End If

            If DDISRULE.SelectedValue <> "" Or DITEMNAME.Text <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DKG.Text <> "" Or DYARD.Text <> "" Then
                If DPRICE.Text = "" Then '檢查 金額是否有資料
                    isOK = False
                    Message = Message + "\n" + "異常：需輸入金額!"
                End If
            End If

            If DITEMNAME.Text <> "" Or DPIECE.Text <> "" Or DMETER.Text <> "" Or DKG.Text <> "" Or DPRICE.Text <> "" Or DYARD.Text <> "" Then
                If DDISRULE.SelectedValue = "" Then '檢查 報廢準則是否有資料
                    isOK = False
                    Message = Message + "\n" + "異常：需輸入報廢準則!"
                End If
            End If


            '  If DITEMNAME.Text <> "" Or DKG.Text <> "" Or DPRICE.Text <> "" Or DDISRULE.SelectedValue <> "" Then
            'If DPIECE.Text = "" And DMETER.Text = "" And DYARD.Text = "" Then
            ' isOK = False
            ' Message = Message + "\n" + "異常：需輸入數量三選一!"
            ' End If
            ' End If



            If Message <> "" Then
                If checkDate() <> 0 Then
                    Message = wRow + "\n" + Message
                End If

                Exit For
            End If

        Next


        If DDISPOSALREASON.SelectedValue = "其它" Then
            If DDISPOSALREASON1.Text = "" Then
                isOK = False
                Message = Message + "\n" + "異常：申請原因選其它，需輸入說明!"
            End If
            If wStep = 1 Then
                DDISPOSALREASON1.BackColor = Color.GreenYellow
            End If

        End If


        If DDUTYDEPO.SelectedValue = "營業" Then
            If DSales.Text = "" Then
                isOK = False
                Message = Message + "\n" + "異常：責任部門是營業，需輸入業務姓名!"
            End If
            If wStep = 1 Then
                DSales.BackColor = Color.GreenYellow
            End If

        End If

        If DSTYPE.SelectedValue = "月次" And DDISPOSALTYPE.Value <> "DEPOT-07" Then
            If DPCNAME.Text = "" Then
                isOK = False
                Message = Message + "\n" + "異常：此申請者無設定扣帳人員" + "\n" + "，請與生管報廢擔當人員確認!"
            End If

        End If

        '判斷是否需立會
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
                Message = Message + "\n" + "異常：請勾選是否需立會!"
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
        '                    Message = Message + "\n" + "異常：請輸入立會時間!"
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
                    Message = Message + "\n" + "上傳格式有誤gridview,請確認!"
                    isOK = False
                End If
                If fielderror = False Then
                    Message = Message + "\n" + "CODE 或 報廢準則 或 部門 不允許空白!"
                    isOK = False
                End If
                If isOK = True Then
                    Insert()
                    If inserterror = False Then
                        Message = Message + "\n" + "上傳格式有誤Insert,請確認!"
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

        If Trim(DDUTYDEPO.SelectedValue) <> "工廠" Then

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
                'DDUTYDEPOName.Text = "張淑敏"
            End If
        Else

            If DMKTSIGN.Checked Then  ' 需營業簽核
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
                    sql = sql + " where cat = 6001 and dkey = 'DUTYDEPO-營業'"
                    sql = sql + " )"
                End If

                Dim dt2 As DataTable = uDataBase.GetDataTable(sql)
                If dt2.Rows.Count > 0 Then
                    DDUTYDEPOID.Text = dt2.Rows(0).Item("userid")
                    DDUTYDEPOName.Text = dt2.Rows(0).Item("username")
                    '  DDUTYDEPOID.Text = "it013"
                    '  DDUTYDEPOName.Text = "張淑敏"
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
                    '   DDUTYDEPOName.Text = "張淑敏"
                End If
            End If
         

        End If
    End Sub

    Protected Sub DDISPOSALREASON_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDISPOSALREASON.SelectedIndexChanged
        DCUSTOMERTOLL.Checked = False
        If DDISPOSALREASON.SelectedValue = "11.其他" Then
            DDISPOSALREASON1.BackColor = Color.GreenYellow
        Else
            DDISPOSALREASON1.BackColor = Color.Yellow
        End If
        If DDUTYDEPO.SelectedValue = "營業" Or DDUTYDEPO.SelectedValue = "營業五課" Then
            DSales.BackColor = Color.GreenYellow
            If DDUTYDEPO.SelectedValue = "營業五課" Then
                DCUSTOMERTOLL.Checked = True
            End If
        Else
            DSales.BackColor = Color.Yellow
        End If
        If DDISPOSALREASON.SelectedValue = "02.緩動品：2年以上未使用" Or DDISPOSALREASON.SelectedValue = "05.2年以上SLD橡膠品" Or DDISPOSALREASON.SelectedValue = "06.在庫大於3年使用量(過剩)" Or DDISPOSALREASON.SelectedValue = "12. 緩動品：1.5年以上未使用" Then
            DDISPOSALRULE.SelectedValue = "低價法對象"
        Else
            DDISPOSALRULE.SelectedValue = ""
        End If

        If DDISPOSALREASON.SelectedValue = "13.尾數/大貨剩餘(FA)" Then
            DSTYPE.SelectedValue = "週次尾數"
        Else
            DSTYPE.SelectedValue = "月次"
        End If

        If wStep = 1 Then  '生管主管KEY立會日期

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

                If DSTYPE.SelectedValue = "週次尾數" Then
                    DSignDate.Text = dtFlow.Rows(0)("lastdate")
                Else

                    DSignDate.Text = dtFlow.Rows(0)("lastdate")

                End If
            Else
                DSignDate.Text = "未開放"

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
            If DDISPOSALTYPE.Value = "完成品" Then
                If DPRICE.Text <> "" Then
                    If CInt(DPRICE.Text) > 100000 Then
                        DREMARK.Text = "合計"
                    Else
                        DREMARK.Text = ""
                    End If
                End If

            End If



            If DDISRULE.SelectedValue = "低價法" Then
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

            ElseIf DDISRULE.SelectedValue = "非低價法" Then
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
        '判斷是否可申請報廢
        Dim SQL, SQL1, Today As String
        Dim SFlag As Integer = 0
        SignDate = ""
        OpenDate = ""
        Today = Now.ToString("yyyy/MM/dd") '現在日時
        ' If wStep = 1 Then
        'Today = Now.ToString("yyyy/MM/dd") '現在日時
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
                '下次開放日
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


    '取得關係人
    Function GetRelated(ByVal userId As String) As String

        Dim sql As String = "select RUserID , RUserName  from M_Related where userid='" & userId & "' and RelatedID='D'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("RUserID")
        End If
        Return NextGate
    End Function

    '重新取得責任部門關係人
    Function GETDUTYDEPOID(ByVal value As String) As String
        Dim UserId As String = ""
        Dim SQL As String
        SQL = " select * from f_disposalsheet"
        SQL = SQL + " where formsno ='" + value.ToString + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
        If Trim(dt.Rows(0)("DutyDepo")) <> "工廠" And Trim(dt.Rows(0)("DutyDepo")) <> "貿易" Then

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

            If Trim(dt.Rows(0)("MKTSIGN")) = 1 Then  ' 需營業簽核
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
                    SQL = SQL + " where cat = 6001 and dkey = 'DUTYDEPO-營業'"
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
        If DDUTYDEPO.SelectedValue = "營業" Or DDUTYDEPO.SelectedValue = "營業五課" Then
            DMKTSIGN.Checked = True
            DSales.BackColor = Color.GreenYellow

            If DDUTYDEPO.SelectedValue = "營業五課" Then
                DCUSTOMERTOLL.Checked = True
            End If

        Else
            DMKTSIGN.Checked = False
            DSales.BackColor = Color.Yellow
        End If
        If DDISPOSALREASON.SelectedValue = "11.其他" Then

            DDISPOSALREASON1.BackColor = Color.GreenYellow
        Else
            DDISPOSALREASON1.BackColor = Color.Yellow
        End If

        uJavaScript.PopMsg(Me, "請再次確認責任部門、是否需勾選營業簽核/向客戶取款")

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

        If DDISPOSALRULE.SelectedValue = "低價法對象" Then
            '細項報廢準則append

            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '6001'"
            sql = sql & " and dkey = 'RULE'"
            sql = sql & " and data = '低價法'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            '新增RULE
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
            sql = sql & " and data = '非低價法'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            '新增RULE
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

            '上傳附檔
            Dim FileName1 As String
            UploadName = DFileUpload1.PostedFile.FileName

            Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                          CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
            Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
            System.Configuration.ConfigurationManager.AppSettings("DISPOSALData")

            '20170912 將檔名修改成不含原始檔名
            Dim fileExtension As String  '副檔名
            fileExtension = IO.Path.GetExtension(UploadName).ToLower   '取得檔案格式
            FileName1 = Path1 + CStr(DNo.Text) + UploadDateTime + fileExtension
            DFileUpload1.PostedFile.SaveAs(FileName1)

            '展開
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
        'uJavaScript.PopMsg(Me, "上傳檔案格式有誤upload,請確認!")
        'griderror = False
        'End Try

        
    End Function

    Sub Insert()
        '檢查是否有同月份同申請人的檔案，若有則將之前的全刪除
        Dim a As String = ""
        Dim uploadflag As Integer = 1
        Dim nullflag As Integer = 1
        Dim sno As String
        Dim NewFormSno As Integer = wFormSno    '表單流水號
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
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        smonth = Mid(DAppDate.Text, 1, 7)


        ' If uploadflag = 1 Then
        Try

            '上傳到資料庫
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
                    If a = "&nbsp;" Then  '檢查第一欄是不是NULL
                        nullflag = 0
                    End If

                Next

                SQL = SQL + jSQL + ","
                SQL = SQL + "getdate(),'" + wApplyID + "')"

                'If CStr(i + 1) = "51" Then
                'uJavaScript.PopMsg(Me, a)
                'End If


                If nullflag = 1 Then '第一欄是空白就不要匯入
                    uDataBase.ExecuteNonQuery(SQL)

                End If


            Next


            uJavaScript.PopMsg(Me, "上傳成功")

            Dim dt_Mail As DataTable
            Dim Adminmail As String = ""
            '表單名稱
            SQL = "select * from M_referp"
            SQL = SQL + " where cat  ='6001'"
            SQL = SQL + " and dkey = 'adminmail'"
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                Adminmail = dt_Mail.Rows(0).Item("Data").ToString
            End If

            '   Send1(wApplyID, Adminmail, wApplyID, "006001", "1", "1", "END")
            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼, 訊息類別

            DInsert.Enabled = False

        Catch ex As Exception
            uJavaScript.PopMsg(Me, "上傳檔案格式有誤Insert,請確認!")
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
                If SheetName = "工作表1$" Or SheetName = "報廢明細$" Then
                Else
                    SheetName = "工作表1$"
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
        '顯示用
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

            '檢查欄位名稱

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
            ElseIf Trim(e.Row.Cells(15).Text.ToUpper) <> "報廢準則" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(16).Text.ToUpper) <> "PN 倉位" Then
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
            ElseIf Trim(e.Row.Cells(23).Text.ToUpper) <> "部門" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(24).Text.ToUpper) <> "營業擔當" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(25).Text.ToUpper) <> "原因" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(26).Text.ToUpper) <> "１年_使用量" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(27).Text.ToUpper) <> "BUYER" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(28).Text.ToUpper) <> "BUYER NAME" Then
                DInsert.Enabled = False
            End If

        'End If




        If DInsert.Enabled = False Then
            Message = "上傳格式有誤gridview,請確認!"
            uJavaScript.PopMsg(Me, Message)
            griderror = False

            GridView1.DataSource = Nothing
            GridView1.DataBind()

        End If

        Else

        '2019/2/14 增加 三個欄位不允許空白
        If Trim(e.Row.Cells(0).Text.ToUpper) <> "&NBSP;" Then
            If Trim(e.Row.Cells(15).Text.ToUpper) = "&NBSP;" Then
                DInsert.Enabled = False
            End If

            If Trim(e.Row.Cells(23).Text.ToUpper) = "&NBSP;" Then
                DInsert.Enabled = False
            End If

        End If

        If DInsert.Enabled = False Then
            Message = "CODE 或 報廢準則 或 部門 不允許空白!"
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
    '** 產生郵件資料至序列檔(Send)
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


        Dim NewFormSno As Integer = wFormSno    '表單流水號
        Dim sno As String
        sno = SetNo(NewFormSno)

        Try



            Dim dt_Mail As DataTable
            '表單名稱
            SQL = "SELECT FormName FROM M_Form "
            SQL &= "Where FormNo  = '" & pFormNo & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xFormName = dt_Mail.Rows(0).Item("FormName").ToString
            End If
            '寄件者郵件帳號,姓名
            SQL = "Select Mail, UserName From M_Users "
            SQL &= "Where UserID  = '" & pFrom & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xFromMail = dt_Mail.Rows(0).Item("Mail").ToString
                xFromName = dt_Mail.Rows(0).Item("UserName").ToString
            End If
            '收件者郵件帳號,姓名
            SQL = "Select Mail, UserName From M_Users "
            SQL &= "Where UserID  = '" & pTo & "' "
            dt_Mail = uDataBase.GetDataTable(SQL)
            If dt_Mail.Rows.Count > 0 Then
                xToMail = dt_Mail.Rows(0).Item("Mail").ToString
                xToName = dt_Mail.Rows(0).Item("UserName").ToString
            End If
            '申請者姓名
            xApplyName = dt_Mail.Rows(0).Item("UserName").ToString
            xStepname = Mid(DAppDate.Text, 1, 6) + "-" + DDepoName.Text + "-" + xFromName + "-" + sno

            '產生郵件資料至序列檔
            SQL = "Insert into Q_WaitSend "
            SQL &= "( "
            SQL &= "Sts, FromID, FromMail, FromName, ToID, "
            SQL &= "ToMail, ToName, CCMail, "
            SQL &= "FormNo, FormSno, FormName, Step, StepName, "
            SQL &= "ApplyID, ApplyName, MSG, MSGName, No, "
            SQL &= "CreateTime "
            SQL &= ") "
            SQL &= "VALUES( "
            SQL &= "'0' ,"                              '控制狀態(0:完成,1:待處理)
            SQL &= "'" & pFrom & "' ,"                  '寄件者ID
            SQL &= "'" & xFromMail & "' ,"              '寄件者郵件帳號
            SQL &= "'" & xFromName & "' ,"              '寄件者姓名
            SQL &= "'" & pTo & "' ,"                    '收件者ID
            SQL &= "'" & xToMail & "' ,"                '收件者郵件帳號
            SQL &= "'" & xToName & "' ,"                '收件者姓名
            SQL &= "'" & xCCMail & "' ,"                'cc者郵件帳號
            SQL &= "'" & pFormNo & "' ,"                '表單No
            SQL &= "'" & CStr(pFormSno) & "' ,"         '表單序號
            SQL &= "'" & xFormName & "' ,"              '表單名稱
            SQL &= "'" & CStr(pStep) & "' ,"            '工程No
            SQL &= "'" & xStepname & "' ,"              '工程名稱
            SQL &= "'" & pApplyID & "' ,"               '申請者ID
            SQL &= "'" & xApplyName & "' ,"             '申請者姓名
            SQL &= "'" & pType & "' ,"                  '訊息類別代碼
            SQL &= "'" & xFlowType & "' ,"              '訊息類別名稱
            SQL &= "'" & xNo & "' ,"                    '委託No
            SQL &= "'" & NowDateTime & "' "             '製作時間
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

        If DDISPOSALRULE.SelectedValue = "低價法對象" Then
            '細項報廢準則append

            sql = "  Select   rank() over(order by [data]) as cno,* from M_referp"
            sql = sql & " where  cat = '6001'"
            sql = sql & " and dkey = 'RULE'"
            sql = sql & " and data = '低價法'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            '新增RULE
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
            sql = sql & " and data = '非低價法'"
            sql = sql & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            '新增RULE
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
        If wStep = 1 Then  '生管主管KEY立會日期

            Dim sql As String

            If DSTYPE.SelectedValue = "週次尾數"   Then
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

                If DSTYPE.SelectedValue = "週次尾數" Then
                    DSignDate.Text = dtFlow.Rows(0)("lastdate")
                    ' DSignDate.Text = "無"
                Else

                    DSignDate.Text = dtFlow.Rows(0)("lastdate")

                End If
            Else
                DSignDate.Text = "未開放"


            End If

            If DSTYPE.SelectedValue = "月次" Then
                '取得扣帳人員
                DPCNAME.Text = GetPCMAN(DAppName.Text)
                If DPCNAME.Text = "" Then
                    uJavaScript.PopMsg(Me, "無設定扣帳人員，請與生管確認")
                End If


            End If

        End If


        

    End Sub

    Function GetPCMAN(ByVal UserName As String) As String

        Dim sql As String

        '= "select  Userid  from M_Users where username=(SELECT SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data)) FROM M_referp where cat ='7002' and Dkey ='RUSER' and Data like '" & UserName & "%') "
        sql = " SELECT * FROM ("
        sql = sql & " SELECT DATA,CASE WHEN SUBSTRING(DATA,1,5) <>'徐萱芳-1' THEN SUBSTRING(DATA,1,CHARINDEX('-',Data)-1)"
        sql = sql & " ELSE SUBSTRING(DATA,1,5) END AS APPNAME"
        sql = sql & " ,CASE WHEN SUBSTRING(DATA,1,5) <>'徐萱芳-1' THEN SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data))"
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
        sql = sql & "  SELECT DATA,CASE WHEN SUBSTRING(DATA,1,6 )not in ('宋美華-PL','宋美華-A1','宋美華-DC') THEN SUBSTRING(DATA,1,CHARINDEX('-',Data)-1) "
        sql = sql & " ELSE SUBSTRING(DATA,1,6) END AS APPNAME "
        sql = sql & " ,CASE WHEN SUBSTRING(DATA,1,6) not in ('宋美華-PL','宋美華-A1','宋美華-DC') THEN SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data))"
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

