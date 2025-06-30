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
    Dim wUserID As String          '使用者ID
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





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
  
        ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查


        ShowFormData()      '顯示表單資料

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()

        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        LFormSno.Text = "單號:" + CStr(wFormSno)
        ' DDISPOSALFILE2.Attributes.Add("onclick", "window.open('file://10.245.1.18/MIS/DASW/D201507001','_blank');return false;")

        '檢查是否已經過實物廢棄完成
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
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()



        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                             System.Configuration.ConfigurationManager.AppSettings("DISPOSALPath")  'WIS-TempPath

        Dim SQL As String
        '主檔資料
        SQL = "Select * From F_DISPOSALSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        '設定檔案路徑

        DMKTSIGN.Checked = False
        DCUSTOMERTOLL.Checked = False

        If DBAdapter1.Rows(0).Item("MKTSIGN") = 1 Then
            DMKTSIGN.Checked = True
        End If

        If DBAdapter1.Rows(0).Item("CUSTOMERTOLL") = 1 Then
            DCUSTOMERTOLL.Checked = True
        End If

        DNo.Text = DBAdapter1.Rows(0).Item("No")

        '開啟報廢資料檔

        '主檔資料
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
            LDISPOSALFILE.NavigateUrl = Path & DBAdapter1.Rows(0).Item("DISPOSALFILE1") '報廢明細
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
            DASDate.Text = DBAdapter1.Rows(0).Item("ASDate")               '立會時間
        End If

        If Mid(DBAdapter1.Rows(0).Item("SignDate").ToString, 1, 4) = "1900" Then
            DSignDate.Text = ""
        Else
            DSignDate.Text = DBAdapter1.Rows(0).Item("SignDate")             '簽核時間
        End If


        '立會時間


        If DBAdapter1.Rows(0).Item("CheckASDate") = "1" Then
            CheckASDate.Items(1).Selected = True
        Else
            CheckASDate.Items(0).Selected = True
        End If



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
            SQL = SQL & " and data ='" + pName + "'"
            SQL = SQL & " order by unique_id"


            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            '新增RULE
            For Each dtr As Data.DataRow In DBAdapter1.Rows


                Rule = "DDISRULE" + Trim(Str(i))
                DText.Items.Add(dtr("Data"))

            Next


        End If




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

