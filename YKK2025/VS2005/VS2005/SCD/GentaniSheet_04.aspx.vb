Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class GentaniSheet_04
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
    Dim wUserIP As String = ""
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
    Dim result As String
 


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()                          '設定參數
        '
        If Not IsPostBack Then                  'PostBack
            SetLinkFile()                       '設定連結檔
            ShowFormData()                      '顯示表單資料
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "CommissionSheet_02.aspx"                                   '程式名
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")                      '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")                            '工程代碼
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
        wFormNo = Request.QueryString("pFormNo")                                                    '表單號碼
        wFormSno = Request.QueryString("pFormSno")                                                  '表單流水號
    End Sub



    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("COMMISSIONFilePath")

        Dim PathOld1 As String = uCommon.GetAppSetting("HttpOldSCD")
        Dim PathOld2 As String = uCommon.GetAppSetting("COMMISSIONFilePathSCD")

        Dim PathOld3 As String = uCommon.GetAppSetting("HttpOldSCD1")
        Dim PathOld4 As String = uCommon.GetAppSetting("NEW-SAMPLE2001FilePath")


        Dim RtnCode As Integer = 0



        'OLD-SCD
        Dim Path_OLDSAMPLE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-SAMPLEFilePath")
        Dim Path_OLDSAMPLEIMAGE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-SAMPLEIMAGEFilePath")
        Dim Path_OLDMANUFACTURE As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-MANUFACTUREFilePath")
        Dim Path_OLDOTHER As String = uCommon.GetAppSetting("HttpOld") & uCommon.GetAppSetting("OLD-OTHERFilePath")
        Dim Path_OLDSAMPLE2001 As String = uCommon.GetAppSetting("HttpOld1") & uCommon.GetAppSetting("OLD-SAMPLE2001FilePath")
        '


        Dim SQL As String

        SQL = "Select * From F_CommissionSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)


        '-----------------------------------------------------------------
        '-- 原單位
        '-----------------------------------------------------------------
        SQL = "Select * From FS_GentaniSheetNEW "
        SQL &= " Where NO =  '" & dtCommissionSheet.Rows(0).Item("NO") & "'"
        Dim dtGentaniSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtGentaniSheet.Rows.Count > 0 Then
            '
            Dim s As New List(Of String)
            s.Add("TNLITEM")
            s.Add("TNRITEM")
            s.Add("ECOL")
            s.Add("EITEM")
            '
            For Each dc As DataColumn In dtGentaniSheet.Columns
                Dim l As Object = Me.form1.FindControl("D4" & dc.ColumnName)
                Dim v As String = uCommon.ReplaceDBnull(dtGentaniSheet.Rows(0)(dc.ColumnName), "")
                If l IsNot Nothing Then
                    l.Text = v
                Else
                    'If s.Contains(dc.ColumnName) Then
                    '    l = Me.form1.FindControl("D4" & dc.ColumnName & "1")
                    '    l.Text = v
                    '    l = Me.form1.FindControl("D4" & dc.ColumnName & "2")
                    '    l.Text = v
                    'End If
                End If
            Next
            D4TNLITEM1.Text = dtGentaniSheet.Rows(0).Item("TNLITEM")
            D4TNRITEM1.Text = dtGentaniSheet.Rows(0).Item("TNRITEM")
            D4ECOL1.Text = dtGentaniSheet.Rows(0).Item("ECOL")
            D4EITEM1.Text = dtGentaniSheet.Rows(0).Item("EITEM")

            '
        End If


    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetLinkFile)
    '**     設定連結檔
    '**
    '*****************************************************************
    Sub SetLinkFile()

     
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
        rqdVal.Style.Add("Top", Top - 50 & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub


    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pIdx As Integer, ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
    End Sub

End Class
