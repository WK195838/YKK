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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            SetParameter()          '設定共用參數
            SetFieldAttribute("")
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
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        Dim SQL As String


        '擔當者及部門 

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DAPPDate.Text = Now.ToString("yyyy/MM/dd") '現在日時
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
        ' DSMONTH.Text = Now.ToString("yyyy/MM") '現在日時
        'DBAdapter1.Clear()

    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        '    oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性

        ' SetFieldAttribute(pPost)     '表單各欄位屬性及欄位輸入檢查等設定
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
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
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator

        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        'rqdVal.Display = ValidatorDisplay.None              ' 因在頁面上加入ValidationSummary , 故驗證控制項統一顯示
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
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".xls", ".xlsx"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9020
            End If
        Next
        'Check上傳檔案Size
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
    '**     判斷是否可繼續執行(驗證資料)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""

        If DSMONTH.Text = "" Then
            isOK = False
            ErrCode = 9010

        End If

        If ErrCode = 0 Then
            If DFileUpload1.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                ErrCode = UPFileIsNormal(DFileUpload1)
            Else
                isOK = False
                ErrCode = 9040
            End If
        End If


        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "請輸入月份"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤9020,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "上傳檔案格式有誤9040,請確認!"
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
        Catch ex As Exception
            uJavaScript.PopMsg(Me, "上傳檔案格式有誤upload,請確認!")

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
        '匯入用
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
        '檢查是否有同月份同申請人的檔案，若有則將之前的全刪除
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
            MsgBoxFlag = MessageBox.Show(DSMONTH.Text + "月份資料已上傳過，是否要將資料清空再重新上傳?", "AlanStudio", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
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
        wApplyID = Request.QueryString("pApplyID")  '申請者ID

        If uploadflag = 1 Then
            Try

                '上傳到資料庫
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
                ' Send1(wApplyID, "pc037", wApplyID, "006002", "1", "1", "END")
                '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別

                DInsert.Enabled = False
            Catch ex As Exception
                uJavaScript.PopMsg(Me, "上傳檔案格式有誤Insert,請確認!")

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
        Str = CStr(DateTime.Now.Year) + Str2
        '年
        'Set單號
        '當月份筆數有幾筆  20150414 Modify by Jessica
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


            '檢查欄位名稱

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
            End If

            'End If




            If DInsert.Enabled = False Then
                Message = "上傳格式有誤gridview,請確認!"
                uJavaScript.PopMsg(Me, Message)

                GridView1.DataSource = Nothing
                GridView1.DataBind()
                GridView2.DataSource = Nothing
                GridView2.DataBind()
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
            xStepname = DSMONTH.Text + "-" + DAPPDepo.Text + "-" + xFromName
          
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

   
End Class

