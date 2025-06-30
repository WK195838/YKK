Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class AddItemList
    Inherits System.Web.UI.Page

    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim remark1, remark2 As String
    Dim DTNo As String


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '*****************************************************************

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        NowDateTime = Now.ToString("yyyyMMddHHmmss")
         
        BExpItemList.Attributes.Add("onClick", "GetExpItem()")
        '
        BADate.Attributes("onclick") = "calendarPicker('Form1.DADate');"


        BReference.Attributes("onclick") = "GetReference();"
        '

        DTNo = Request.QueryString("pDTNo")
        wStep = Request.QueryString("pStep")
        If wStep = 20 Or wStep = 40 Then
            BEDIT.Visible = True
        Else
            BEDIT.Visible = False
        End If

        If Not Me.IsPostBack Then   '不是PostBack
            SetFieldAttribute("New")

            SetFieldData()

            'If Len(DTNo) <> 14 Then
            Dim sql As String
            sql = " delete from F_FundingSheetdttemp  where dtno ='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(sql)

            sql = " insert into F_FundingSheetdttemp (dtno,expcat,expitem,acid,adate,taxtype,netamt,taxamt,amt,content,remark,InvoiceNo,TaxNo,GUINo,TaxBase,createuser,createtime)"
            sql = sql + " Select  no,expcat,expitem,acid,adate,taxtype,netamt,taxamt,amt,content,remark,InvoiceNo,TaxNo,GUINo,TaxBase,createuser,createtime  from  F_FundingSheetdt "
            sql = sql + " where no ='" + DTNo + "' order by unique_id"
            uDataBase.ExecuteNonQuery(sql)
            'End If

        Else


        End If


       

        GetData()
 

        '
    End Sub


    Protected Sub BADD_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADD.ServerClick
        If InputDataOK(0) Then
            Dim ACID As String
            Dim sql As String = ""
            ACID = DACID.Text

            If DTaxType.SelectedItem.Text <> "繳納證(V8)" Then
                DTaxNo.Text = ""
            Else
                DInvoiceNo.Text = ""
                DGUINo.Text = ""
            End If


            sql = " Insert into F_FundingSheetdtTemp (DTNo, ExpCat, ExpItem,ACID, ADate, TaxType, NetAmt, TaxAmt, Amt, Content, Remark,InvoiceNo,TaxNo,GUINo,TaxBase,Createuser,CreateTime,ModifyUser,ModifyTime) "
            sql = sql & "VALUES( "
            sql = sql & " '" & DTNo & "', "
            sql = sql & " '" & Mid(DExpItem.Text, 1, InStr(DExpItem.Text, "--") - 1) & "', "
            sql = sql & " '" & Mid(DExpItem.Text, InStr(DExpItem.Text, "--") + 2) & "', "
            sql = sql & " '" & DACID.Text & "', "
            sql = sql & " '" & DADate.Text & "', "
            sql = sql & " '" & DTaxType.SelectedItem.Text & "', "
            sql = sql & " " & Replace(DNetAmt.Text, ",", "") & ", "
            sql = sql & " " & Replace(DTaxAmt.Text, ",", "") & ", "
            sql = sql & " " & Replace(DAmt.Text, ",", "") & ", "
            sql = sql & " N'" & DContent.Text & "', "
            sql = sql & " N'" & DRemark.Text & "', "
            sql = sql & "'" & UCase(DInvoiceNo.Text) & "', "
            sql = sql & "'" & UCase(DTaxNo.Text) & "', "
            sql = sql & "'" & DGUINo.Text & "', "
            sql = sql & " " & Replace(DTaxBase.Text, ",", "") & ", "
            sql = sql & " '" & Request.QueryString("pUserID") & "', "
            sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
            sql = sql & " '" & Request.QueryString("pUserID") & "', "
            sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
            uDataBase.ExecuteNonQuery(sql)


            '
            GetData()

            '2025 登錄後清空
            DADate.Text = ""
            DInvoiceNo.Text = ""
            DTaxNo.Text = ""
            DGUINo.Text = ""
            DAmt.Text = "0"
            DNetAmt.Text = "0"
            DTaxAmt.Text = "0"
            DTaxBase.Text = "0"
            DExpItem.Text = ""
            DContent.Text = ""
            DRemark.Text = ""
            DTaxType.Text = ""


        End If


    End Sub

    Sub GetData()
        Dim SQL As String

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定


  
        SQL = " Select  Unique_ID,ExpCat+'--'+ExpItem ExpItem,ACID, ADate, TaxType, NetAmt, TaxAmt, Amt, Content, Remark,case when InvoiceNo ='' then Taxno else InvoiceNo end as InvoiceNo,GUINo,TaxBase from  F_FundingSheetdttemp  where 1=1 "
        SQL = SQL + " and DTNO = '" + DTNo + "'"
            SQL = SQL + " order by unique_id "


     
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Getata")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
        If wStep = 20 Or wStep = 40 Then
            GridView1.AutoGenerateSelectButton = True
        Else
            GridView1.AutoGenerateSelectButton = False
        End If


        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(1).Visible = False
        'If wStep = 20 Or wStep = 40 Then
        '    GridView1.AutoGenerateSelectButton = True
        'Else
        '    GridView1.AutoGenerateSelectButton = False
        'End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        Dim id As String = GridView1.DataKeys(e.RowIndex).Value

        SQL = " delete  from  F_FundingSheetdtTemp "
        SQL = SQL & " where unique_id = " & id & " "
        uDataBase.ExecuteNonQuery(SQL)
        '
        GetData()
    End Sub

    Protected Sub DAmt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAmt.TextChanged
        Dim TaxP As Decimal = 0

        TaxP = 0.05


        If CLng(DAmt.Text) <> 0 Then
            If InStr(DTaxType.SelectedValue, "V1") > 0 Then
                DTaxAmt.Text = "0"
                DNetAmt.Text = DAmt.Text

            Else
                'DTaxAmt.Text = CLng(CLng(DAmt.Text) * TaxP + 0.5)
                'DNetAmt.Text = CLng(CLng(DAmt.Text) - CLng(DTaxAmt.Text))
                DNetAmt.Text = Math.Round(CLng(DAmt.Text) / (1 + TaxP), 0)
                DTaxAmt.Text = CLng(DAmt.Text) - CLng(DNetAmt.Text)

            End If

            '
            DAmt.Text = String.Format("{0:N0}", Val(CLng(DAmt.Text)))
            DTaxAmt.Text = String.Format("{0:N0}", Val(CLng(DTaxAmt.Text)))
            DNetAmt.Text = String.Format("{0:N0}", Val(CLng(DNetAmt.Text)))
            If DTaxType.SelectedValue = "收據(V1)" Then
                DTaxBase.Text = "0"
            Else
                DTaxBase.Text = DNetAmt.Text
            End If
        End If
    End Sub

    Protected Sub DTaxType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTaxType.SelectedIndexChanged
        '2025
        '   DExpItem.Text = ""
        '   DContent.Text = ""
        '  DInvoiceNo.Text = ""
        '  DTaxNo.Text = ""
        '  DGUINo.Text = ""
        ' DContent1.Text = ""

        If DTaxType.Text <> "" Then
            BExpItemList.Disabled = False
            BContent.Disabled = False
            BRemark.Disabled = False
            DInvoiceNo.Visible = True

            '2024/10/29 jessica 
            If DTaxType.SelectedValue = "收據(V1)" Then
                DInvoiceNo.BackColor = Color.Yellow
                DGUINo.BackColor = Color.Yellow
                DTaxNo.Visible = False
                ' DTaxNo.Text = ""
                DGUINo.Visible = True
            ElseIf DTaxType.SelectedValue = "繳納證(V8)" Then
                DInvoiceNo.Visible = False
                'DInvoiceNo.Text = ""
                DGUINo.Visible = False
                'DGUINo.Text = ""
                DTaxNo.BackColor = Color.GreenYellow
                DTaxNo.Visible = True
            Else
                DInvoiceNo.BackColor = Color.GreenYellow
                DGUINo.BackColor = Color.GreenYellow
                DTaxNo.Visible = False
                DGUINo.Visible = True
                'DTaxNo.Text = ""
                'DInvoiceNo.Text = ""
                'DGUINo.Text = ""
            End If
            '2024/10/29 jessica 
      
        Else
            BExpItemList.Disabled = True
            BContent.Disabled = True
            BRemark.Disabled = True
         
        End If
        Dim TaxP As Decimal = 0
        '  If DTaxType.SelectedValue <> "" Then
        TaxP = 0.05
        'End If


        If CLng(DAmt.Text) <> 0 Then

            If InStr(DTaxType.SelectedValue, "V1") > 0 Then
                DTaxAmt.Text = "0"
                DNetAmt.Text = CLng(DAmt.Text)
            Else
                ' DTaxAmt.Text = CLng(CLng(DAmt.Text) * TaxP + 0.5)
                'DNetAmt.Text = CLng(CLng(DAmt.Text) - CLng(DTaxAmt.Text))
                '2025
                DNetAmt.Text = Math.Round(CLng(DAmt.Text) / (1 + TaxP), 0)
                DTaxAmt.Text = CLng(DAmt.Text) - CLng(DNetAmt.Text)

            End If
            '
            DAmt.Text = String.Format("{0:N0}", Val(CLng(DAmt.Text)))
            DTaxAmt.Text = String.Format("{0:N0}", Val(CLng(DTaxAmt.Text)))
            DNetAmt.Text = String.Format("{0:N0}", Val(CLng(DNetAmt.Text)))
            If DTaxType.SelectedValue = "收據(V1)" Then
                DTaxBase.Text = "0"
            Else
                DTaxBase.Text = DNetAmt.Text
            End If

        End If


    End Sub

    Protected Sub BClose_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.ServerClick
        Dim sql As String




        'jessica 2021/2/5
        sql = "Select * From F_FundingSheetdtTemp "
        sql = sql & "Where DTNO = '" + DTNo + "' "
        Dim dtTemp As DataTable = uDataBase.GetDataTable(sql)
        If dtTemp.Rows.Count > 0 Then
            '將登錄的細項INERT 到DT 
            sql = " delete from F_FundingSheetdt  where no ='" + DTNo + "'"
            uDataBase.ExecuteNonQuery(sql)

            sql = " insert into F_FundingSheetdt (no,expcat,expitem,acid,adate,taxtype,netamt,taxamt,amt,content,remark,InvoiceNo,TaxNo,GUINo,TaxBase,createuser,createtime)"
            sql = sql + " Select  dtno,expcat,expitem,acid,adate,taxtype,netamt,taxamt,amt,content,remark,InvoiceNo,TaxNo,GUINo,TaxBase,createuser,createtime  from  F_FundingSheetdttemp "
            sql = sql + " where dtno ='" + DTNo + "' order by unique_id"
            uDataBase.ExecuteNonQuery(sql)
        End If



        Dim js As String = ""
        js &= "window.opener.document.all.DUpdate.value = '1';"
        ' js &= "window.close();"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        '
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub DExpItem_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DExpItem.TextChanged

        'remark1 = "1." + Mid(DExpItem.Text, InStr(4, DExpItem.Text, "-", 1) + 1, Len(DExpItem.Text) - 1)
        'remark2 = "2." + Mid(DExpItem.Text, InStr(4, DExpItem.Text, "-", 1) + 1, Len(DExpItem.Text) - 1)
        remark1 = "1." + DExpItem.Text
        remark2 = "2." + DExpItem.Text
        Dim TaxType As String
        TaxType = DTaxType.SelectedValue
        BContent.Attributes.Add("onClick", "GetRemark('" + remark1 + "," + TaxType + "')") '
        BRemark.Attributes.Add("onClick", "GetRemark('" + remark2 + "')") '

    End Sub


    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        DADate.Visible = True
        DADate.BackColor = Color.GreenYellow
        DADate.Attributes.Add("readonly", "true")
   
        'ShowRequiredFieldValidator("DADateRqd", "DADate", "異常：需輸入日期")

        DAmt.Visible = True
        DAmt.BackColor = Color.GreenYellow
        DAmt.ReadOnly = False
        'ShowRequiredFieldValidator("DAMTRqd", "DAMT", "異常：需輸入總額")


        DTaxType.Visible = True
        DTaxType.BackColor = Color.GreenYellow
        'ShowRequiredFieldValidator("DAMTRqd", "DAMT", "異常：需輸入稅別")

        DTaxAmt.Visible = True
        DTaxAmt.BackColor = Color.GreenYellow
        DTaxAmt.ReadOnly = False
        'ShowRequiredFieldValidator("DTaxAMTRqd", "DTaxAMT", "異常：需輸入稅額")

        DNetAmt.Visible = True
        DNetAmt.BackColor = Color.GreenYellow
        DNetAmt.ReadOnly = False
        'ShowRequiredFieldValidator("DNetAmtRqd", "DNetAmt", "異常：需輸入淨額")

        If wStep <> 40 Then
            DTaxBase.Visible = True
            DTaxBase.BackColor = Color.LightGray
            DTaxBase.ReadOnly = True
        Else
            DTaxBase.Visible = True
            DTaxBase.BackColor = Color.GreenYellow
            DTaxBase.ReadOnly = False
        End If



        DExpItem.Visible = True
        DExpItem.BackColor = Color.GreenYellow
        DExpItem.ReadOnly = True
        'ShowRequiredFieldValidator("DExpItemRqd", "DExpItem", "異常：需輸入類別")



        DContent.Visible = True
        DContent.BackColor = Color.GreenYellow
        DContent.ReadOnly = True

        'ShowRequiredFieldValidator("DContentRqd", "DContent", "異常：需輸入內容")


        DRemark.Visible = True
        DRemark.BackColor = Color.Yellow
        DRemark.ReadOnly = True
        'ShowRequiredFieldValidator("DC
    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()

        Dim SQL As String
        Dim i As Integer
        SQL = "  Select taxname+'('+taxid+')' as TaxName,TaxP from m_taxList order by taxname  "

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DTaxType.Items.Clear()
        DTaxType.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("TaxName")
            ListItem1.Value = dtReferp.Rows(i).Item("TaxName")
            DTaxType.Items.Add(ListItem1)

        Next
        dtReferp.Clear()

    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top + 20 & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub

    Function InputDataOK(ByVal pAction As Integer) As Boolean

        Dim Message As String = ""
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False

        If ErrCode = 0 Then
            If DADate.Text = "" Then
                ErrCode = 1

            End If
        End If

        If ErrCode = 0 Then

            If DAmt.Text = "0" Then
                ErrCode = 2
            End If

        End If




        If ErrCode = 0 Then

            If DNetAmt.Text = "0" Then
                'ErrCode = 4
            End If
        End If

        If ErrCode = 0 Then

            If DExpItem.Text = "" Then
                ErrCode = 5
            End If
        End If

        If ErrCode = 0 Then

            If DContent.Text = "" Then
                ErrCode = 6
            End If
        End If


        If ErrCode = 0 Then

            If DTaxType.Text = "" Then
                ErrCode = 7
            End If

        End If


        '2024/10/29 jessica 
        If ErrCode = 0 Then
            If DTaxType.SelectedValue <> "收據(V1)" Then

                If DTaxType.SelectedValue = "繳納證(V8)" Then
                    '稅單號碼檢查
                    If DTaxNo.Text = "" Then
                        ErrCode = 8
                    End If

                    If IsEnglish(UCase(Mid(DTaxNo.Text, 1, 3))) = False Then
                        ErrCode = 12
                    End If
                    If UCase(Mid(DTaxNo.Text, 3, 1)) <> "I" Then
                        ErrCode = 12
                    End If
                    '2025
                    '     If IsInt(Mid(DTaxNo.Text, 4, 11)) = False Then
                    'ErrCode = 12
                    'End If
                    '2025
                    If Len(DTaxNo.Text) <> 14 Then
                        ErrCode = 12
                    End If


                    '稅單號碼檢查
                Else
                    If DInvoiceNo.Text = "" Then
                        ErrCode = 9
                    Else
                        '發票號碼檢查
                        If Len(DInvoiceNo.Text) <> 10 Then
                            ErrCode = 11
                        Else

                            If IsEnglish(UCase(Mid(DInvoiceNo.Text, 1, 2))) = False Then
                                ErrCode = 11
                            End If
                            If IsInt(Mid(DInvoiceNo.Text, 3, 8)) = False Then
                                ErrCode = 11
                            End If

                        End If
                        '發票號碼檢查


                    End If

                End If


            End If


        End If

        If ErrCode = 0 Then
            If CInt(DAmt.Text) <> CInt(DTaxAmt.Text) + CInt(DNetAmt.Text) Then
                ErrCode = 13
            End If
        End If

        '2025
        If ErrCode = 0 Then
            If DTaxType.SelectedValue <> "繳納證(V8)" And DTaxType.SelectedValue <> "收據(V1)" Then
                If DGUINo.Text <> "" Then
                    If Len(DGUINo.Text) <> 8 Then
                        ErrCode = 10
                    End If
                Else
                    ErrCode = 14
                End If

            End If
        End If

     



        '2024/10/29 jessica 

        If ErrCode = 1 Then Message = "異常：需輸入日期"
        If ErrCode = 2 Then Message = "異常：需輸入總額"
        If ErrCode = 3 Then Message = "異常：需輸入稅額"
        If ErrCode = 4 Then Message = "異常：需輸入淨額"
        If ErrCode = 5 Then Message = "異常：需輸入類別"
        If ErrCode = 6 Then Message = "異常：需輸入內容"
        If ErrCode = 7 Then Message = "異常：需輸入稅別"
        If ErrCode = 8 Then Message = "異常：需輸入稅單號碼"
        If ErrCode = 9 Then Message = "異常：需輸入發票號碼"
        If ErrCode = 10 Then Message = "異常：賣方統一編號輸入錯誤"
        If ErrCode = 11 Then Message = "異常：發票號碼輸入錯誤"
        If ErrCode = 12 Then Message = "異常：稅單號碼輸入錯誤"
        If ErrCode = 13 Then Message = "異常：總額有誤，請再確認"
        If ErrCode = 14 Then Message = "異常：需輸入賣方統一編號"




        If ErrCode <> 0 Then
            isOK = False
            uJavaScript.PopMsg(Me, Message)

        Else
            isOK = True
        End If

        Return isOK


    End Function

    Protected Sub DExpItem1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DExpItem1.TextChanged

        DContent.Text = ""
        DRemark.Text = ""
        DContent1.Text = ""
        Dim TaxType As String
        TaxType = DTaxType.SelectedValue


        DExpItem.Text = DExpItem1.Text
        remark1 = "1." + DExpItem.Text
        remark2 = "2." + DExpItem.Text
        '  BContent.Attributes.Add("onClick", "GetRemark('" + remark1 + "')") '
        BContent.Attributes.Add("onClick", "GetRemark('" + remark1 + "','" + TaxType + "')") '
        BRemark.Attributes.Add("onClick", "GetRemark('" + remark2 + "')") '





    End Sub


    Protected Sub DContent1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DContent1.TextChanged
        DContent.Text = DContent1.Text

    End Sub

    Protected Sub DRemark1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DRemark1.TextChanged
        DRemark.Text = DRemark1.Text
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim SQL As String

        Dim id As String = GridView1.SelectedValue

        SQL = " select* from  F_FundingSheetdttemp "
        SQL = SQL & " where unique_id = " & id & " "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        '
        DADate.Text = dtData.Rows(0).Item("ADate")
        DTaxType.Text = dtData.Rows(0).Item("TaxType")
        DACID.Text = dtData.Rows(0).Item("ACID")
        DAmt.Text = dtData.Rows(0).Item("Amt")
        DTaxAmt.Text = dtData.Rows(0).Item("TaxAmt")
        DNetAmt.Text = dtData.Rows(0).Item("NetAmt")
        DExpItem.Text = dtData.Rows(0).Item("ExpCat") + "--" + dtData.Rows(0).Item("ExpItem")
        DContent.Text = dtData.Rows(0).Item("Content")
        DRemark.Text = dtData.Rows(0).Item("Remark")
        DInvoiceNo.Text = dtData.Rows(0).Item("InvoiceNo")
        DTaxNo.Text = dtData.Rows(0).Item("TaxNo")
        DGUINo.Text = dtData.Rows(0).Item("GUINo")
        DTaxBase.Text = dtData.Rows(0).Item("TaxBase")

        DID.Text = id
        '2005 
        BExpItemList.Disabled = False
        BContent.Disabled = False
        BRemark.Disabled = False

        remark1 = "1." + DExpItem.Text
        remark2 = "2." + DExpItem.Text
        Dim TaxType As String
        TaxType = DTaxType.SelectedValue

        BContent.Attributes.Add("onClick", "GetRemark('" + remark1 + "','" + TaxType + "')") '
        BRemark.Attributes.Add("onClick", "GetRemark('" + remark2 + "')")



    End Sub

    Protected Sub BEDIT_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEDIT.ServerClick
        Dim Message As String = ""


        '2024/10/29 jessica 
        '2025
        If DTaxType.SelectedItem.Text <> "繳納證(V8)" Then
            DTaxNo.Text = ""
        Else
            DInvoiceNo.Text = ""
            DGUINo.Text = ""
        End If

    
        Dim SQL As String
        Try

            If DID.Text <> "" Then
                If InputDataOK(0) Then
                    SQL = " delete  from  F_FundingSheetdttemp "
                    SQL = SQL & " where unique_id = " & DID.Text & " "
                    uDataBase.ExecuteNonQuery(SQL)

                    Dim ACID As String

                    ACID = DACID.Text

                    SQL = " Insert into F_FundingSheetdtTemp (DTNo, ExpCat, ExpItem,ACID, ADate, TaxType, NetAmt, TaxAmt, Amt, Content, Remark,InvoiceNo,TaxNo,GUINo,TaxBase,Createuser,CreateTime,ModifyUser,ModifyTime) "
                    SQL = SQL & "VALUES( "
                    SQL = SQL & " '" & DTNo & "', "
                    SQL = SQL & " '" & Mid(DExpItem.Text, 1, InStr(DExpItem.Text, "--") - 1) & "', "
                    SQL = SQL & " '" & Mid(DExpItem.Text, InStr(DExpItem.Text, "--") + 2) & "', "
                    SQL = SQL & " '" & DACID.Text & "', "
                    SQL = SQL & " '" & DADate.Text & "', "
                    SQL = SQL & " '" & DTaxType.SelectedItem.Text & "', "
                    SQL = SQL & " " & Replace(DNetAmt.Text, ",", "") & ", "
                    SQL = SQL & " " & Replace(DTaxAmt.Text, ",", "") & ", "
                    SQL = SQL & " " & Replace(DAmt.Text, ",", "") & ", "
                    SQL = SQL & " N'" & DContent.Text & "', "
                    SQL = SQL & " N'" & DRemark.Text & "', "
                    SQL = SQL & "'" & UCase(DInvoiceNo.Text) & "', "
                    SQL = SQL & "'" & UCase(DTaxNo.Text) & "', "
                    SQL = SQL & "'" & DGUINo.Text & "', "
                    SQL = SQL & " " & Replace(DTaxBase.Text, ",", "") & ", "
                    SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
                    SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
                    SQL = SQL & " '" & Request.QueryString("pUserID") & "', "
                    SQL = SQL & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
                    uDataBase.ExecuteNonQuery(SQL)



                    Message = "資料已修改!"
                End If
            Else
                Message = "未選取資料!"
            End If


        Catch ex As Exception
            Message = "異常,請連絡系統人員!"

        End Try


        uJavaScript.PopMsg(Me, Message)

        '恢復未選取
        DID.Text = ""
        GridView1.SelectedIndex = -1
        GetData()
        'DExpItem.Text = ""
        'DContent.Text = ""
        'DInvoiceNo.Text = ""
        'DTaxNo.Text = ""
        'DGUINo.Text = ""
        'DContent1.Text = ""
        'DAmt.Text = ""
        'DTaxAmt.Text = ""
        'DNetAmt.Text = ""
        'DTaxBase.Text = ""



    End Sub

    '判斷是否為英文字   '2024/10/29 jessica 

    Private Function IsEnglish(ByVal str As String) As Boolean
        Dim eng As Regex = New Regex("^[A-Z]+$")
        Return eng.IsMatch(str)
    End Function


    '判斷是否為數字   '2024/10/29 jessica 
    Public Shared Function IsInt(ByVal str As String) As Boolean
        Dim eng As Regex = New Regex("^[0-9]*[1-9][0-9]*$")
        Return eng.IsMatch(Str)
    End Function
 
    Protected Sub DNetAmt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNetAmt.TextChanged
        '2025
        If DTaxType.SelectedValue = "收據(V1)" Then
            DTaxBase.Text = "0"
        Else

            DTaxBase.Text = DNetAmt.Text
        End If

    End Sub
 
End Class
