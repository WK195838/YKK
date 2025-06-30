Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb


 


Partial Class DTMW_NewColorinqCommission
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "DTMW_inqCommission.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '現在日時
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub


    Sub SetSearchItem()
        Dim SQL As String
        Dim I As Integer
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '5001'"
        SQL = SQL & " and dkey = 'YKKColorType'"


        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DYKKColortype.Items.Add("")
        For I = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(I).Item("Data")
            ListItem1.Value = dtReferp.Rows(I).Item("Data")
            DYKKColortype.Items.Add(ListItem1)
        Next

        BASDate.Attributes("onclick") = "calendarPicker('Form1.DASDate');"
        BAEDate.Attributes("onclick") = "calendarPicker('Form1.DAEDate');"
 
        BFSDate.Attributes("onclick") = "calendarPicker('Form1.DFSDate');"
        BFEDate.Attributes("onclick") = "calendarPicker('Form1.DFEDate');"

        BAFSDate.Attributes("onclick") = "calendarPicker('Form1.DAFSDate');"
        BAFEDate.Attributes("onclick") = "calendarPicker('Form1.DAFEDate');"

        BSts.Attributes("onclick") = "GetStsPicker();"
        BFormNo.Attributes("onclick") = "GetFormPicker();"
        BDivision.Attributes("onclick") = "GetDivisionPicker();"

    End Sub


    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '程式別不同

        Dim path As String
        path = "D:\db\DTMW_NewColorCommission_ist.xls"
        ' pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=" + path)     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        If DFormNo2.Text = "005012" Then
            Datagrid2.RenderControl(hw)
            Datagrid2.AllowPaging = wAllowPaging        '程式別不同
        Else
            DataGrid1.RenderControl(hw)
            DataGrid1.AllowPaging = wAllowPaging        '程式別不同
        End If



        Response.Write(tw.ToString())
        Response.End()


    End Sub


    Sub DataList()
        Dim ASdate As Date
        Dim ASdate1 As String = ""
        If DASDate.Text <> "" Then
            ASdate = DASDate.Text
            ASdate1 = Format(ASdate, "yyyy/MM/dd")
            ASdate1 = " and  Convert(VARCHAR(10), ApplyTime, 111) >= '" + ASdate1 + "'"
        End If


        Dim AEdate As Date
        Dim AEdate1 As String = ""

        If DAEDate.Text <> "" Then
            AEdate = DAEDate.Text
            AEdate1 = Format(AEdate, "yyyy/MM/dd")
            AEdate1 = " and  Convert(VARCHAR(10), ApplyTime, 111) <= '" + AEdate1 + "'"
        End If

        Dim FSdate As Date
        Dim FSdate1 As String = ""
        If DFSDate.Text <> "" Then
            FSdate = DFSDate.Text
            FSdate1 = Format(FSdate, "yyyy/MM/dd")
            FSdate1 = " and DeliveryDate >= '" + FSdate1 + "'"
        End If

        Dim FEdate As Date
        Dim FEdate1 As String = ""

        If DFEDate.Text <> "" Then
            FEdate = DFEDate.Text
            FEdate1 = Format(FEdate, "yyyy/MM/dd")
            FEdate1 = " and DeliveryDate <= '" + FEdate1 + "'"
        End If

        Dim AFSdate As Date
        Dim AFSdate1 As String = ""
        If DAFSDate.Text <> "" Then
            AFSdate = DAFSDate.Text
            AFSdate1 = Format(AFSdate, "yyyy/MM/dd")
            AFSdate1 = " and completedDate >= '" + AFSdate1 + "'"
        End If


        Dim AFEdate As Date
        Dim AFEdate1 As String = ""

        If DAFEDate.Text <> "" Then
            AFEdate = DAFEDate.Text
            AFEdate1 = Format(AFEdate, "yyyy/MM/dd")
            AFEdate1 = " and CompletedDate <= '" + AFEdate1 + "'"
        End If




        Dim wSts As String = ""
        wSts = DSts2.Text
        If DSts1.Text <> "" Then
            wSts = " and b.Sts in (" + wSts + ")"
        End If

        Dim wFormNo As String = ""
        wFormNo = DFormNo2.Text
        If DFormNo1.Text <> "" Then
            wFormNo = " and a.FormNo in (" + wFormNo + ")"
        Else
            wFormNo = ""
        End If

        If DFormNo1.Text = "型別轉換-05CNPBS16" Then
            ASdate1 = " and  Convert(VARCHAR(10), ApplyTime, 111) >= '2025/04/01' "
        End If

        Dim wDivision As String = ""
        wDivision = DDivision2.Text
        If DDivision1.Text <> "" Then
            wDivision = " and left(a.hrwdivid,7) in (" + wDivision + ")"
        Else
            wDivision = ""
        End If

        Dim wNo As String = ""
        If DNo.Text <> "" Then
            wNo = " and  b.No like '%" + DNo.Text + "%'"
        Else
            wNo = ""
        End If

        Dim wName As String = ""
        If DName.Text <> "" Then
            wName = " and  applyName like '%" + DName.Text + "%'"
        Else
            wName = ""
        End If

        Dim wBuyer As String = ""
        If DBuyer.Text <> "" Then
            wBuyer = " and b.buyer like '%" + DBuyer.Text + "%'"
        Else
            wBuyer = ""

        End If


        Dim wCustomer As String = ""
        If DCustomer.Text <> "" Then
            wCustomer = " and b.customer like '%" + DCustomer.Text + "%'"
        Else
            wCustomer = ""
        End If

        Dim wYKKColorType As String = ""
        If DYKKColortype.SelectedValue <> "" Then
            wYKKColorType = " and YKKcolortype like '%" + DYKKColortype.SelectedValue + "%'"
        Else
            wYKKColorType = ""
        End If

        Dim wYKKColorcode As String = ""
        If DYKKColorCode.Text <> "" Then
            wYKKColorcode = " and YKKcolorCode like '%" + DYKKColorCode.Text + "%'"
        Else
            wYKKColorcode = ""
        End If

        Dim wSLDColor As String = ""
        If DSLDColor.Text <> "" Then
            wSLDColor = " and SLDColor like '%" + DSLDColor.Text + "%'"
        Else
            wSLDColor = ""
        End If

        Dim wVFColor As String = ""
        If DVFColor.Text <> "" Then
            wVFColor = " and VFColor like '%" + DVFColor.Text + "%'"
        Else
            wVFColor = ""
        End If

        Dim wNewOldColor As String = ""

        If DNewOldColor.SelectedValue <> "" Then
            wNewOldColor = " and NewOldColor = '" + DNewOldColor.SelectedValue + "'"
        Else
            wNewOldColor = ""
        End If

        Dim wCustomerColor As String = ""
        If DCustomerColor.Text <> "" Then
            wCustomerColor = " and CustomerColor like '%" + DCustomerColor.Text + "%'"
        Else
            wCustomerColor = ""
        End If


        Dim wCustomerColorCode As String = ""
        If DCustomerColorCode.Text <> "" Then
            wCustomerColorCode = " and CustomerColorCode like '%" + DCustomerColorCode.Text + "%'"
        Else
            wCustomerColorCode = ""
        End If

        Dim wOverSeaYkkcode As String = ""
        If DOverSeaYkkcode.Text <> "" Then
            wOverSeaYkkcode = " and OverSeaYkkcode like '%" + DOverSeaYkkcode.Text + "%'"
        Else
            wOverSeaYkkcode = ""
        End If

        Dim wPANTONECode As String = ""
        If DPANTONECode.Text <> "" Then
            wPANTONECode = " and PANTONECode like '%" + DPANTONECode.Text + "%'"
        Else
            wPANTONECode = ""
        End If



        Dim wReColorCode As String = ""
        If DReColorCode.Text <> "" Then
            wReColorCode = " and ReColorCode like '%" + DReColorCode.Text + "%'"
        Else
            wReColorCode = ""
        End If

        Dim TABLE As String = ""
        If DTABLE.SelectedValue = "未封存" Then
            TABLE = "V_WaitHandle_01"
        Else
            TABLE = "V_WaitHandle_old_01"
        End If


        Dim i As Integer = 0
        Dim SQL As String
        If DFormNo2.Text <> "005012" Then
            DataGrid1.Columns.Item(0).HeaderText = "編號"
            DataGrid1.Columns.Item(1).HeaderText = "依賴狀態"
            DataGrid1.Columns.Item(2).HeaderText = "依賴日期"
            DataGrid1.Columns.Item(3).HeaderText = "完成日"

            DataGrid1.Columns.Item(4).HeaderText = "依賴表單"
            DataGrid1.Columns.Item(5).HeaderText = "依賴部門"
            DataGrid1.Columns.Item(6).HeaderText = "依賴者"
            DataGrid1.Columns.Item(7).HeaderText = "客戶"
            DataGrid1.Columns.Item(8).HeaderText = "BUYER"

            DataGrid1.Columns.Item(9).HeaderText = "客戶色名"
            DataGrid1.Columns.Item(10).HeaderText = "客戶色號"
            DataGrid1.Columns.Item(11).HeaderText = "海外YKK色號"
            DataGrid1.Columns.Item(12).HeaderText = "PANTONE色號"

            DataGrid1.Columns.Item(13).HeaderText = "YKK色別"
            DataGrid1.Columns.Item(14).HeaderText = "YKK色號"
            DataGrid1.Columns.Item(15).HeaderText = "兼用色拉頭"


            DataGrid1.Columns.Item(16).HeaderText = "兼用色VF上/下止"

            DataGrid1.Columns.Item(17).HeaderText = "兼用色VF霧齒"
            DataGrid1.Columns.Item(18).HeaderText = "新舊色"

            If DFormNo2.Text = "005013" Then
                DataGrid1.Columns.Item(19).HeaderText = "回收紗色號"

            Else
                DataGrid1.Columns.Item(19).HeaderText = "對應YKK色號"

            End If




            SQL = "SELECT "
            SQL = SQL + " b.No  As Field1,"
            SQL = SQL + " case b.Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2, "
            SQL = SQL + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3, "
            SQL = SQL + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate, "

            SQL = SQL + " case when a.formno ='005008' and convert(char(10),getdate(),111) >='2025/04/01' then '型別轉換-05CNPBS16' else  a.FormName  end as Field4,"
            SQL = SQL + " a.Divname As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"

            SQL = SQL + " YKKColorType As Field9,YKKColorCode as Field10, case  when a.formno ='005011' then  YKKColorCodeSLD else sldcolor end As Field11,"
            SQL = SQL + " case when a.formno ='005011' then  VFCOLOR else VFCOLOR end As Field12,case when a.formno ='005011' then ykkcolorcodevf else vfcolor end as Field13,NewOldColor,"


            SQL = SQL + " '....' as WorkFlow, ViewURL, "

            If DTABLE.SelectedValue = "未封存" Then
                SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
                SQL = SQL + "'pFormNo='   + a.FormNo + "
                SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
                SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
                SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
                SQL = SQL + "'&pApplyID=' + a.ApplyID "
                SQL = SQL + " As OPURL,  "
            Else
                SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
                SQL = SQL + "'pFormNo='   + a.FormNo + "
                SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
                SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
                SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
                SQL = SQL + "'&pApplyID=' + a.ApplyID + '&pKeepdata=1'"
                SQL = SQL + " As OPURL,  "
            End If

            SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,customerColorCode,overSeaYkkCode,pantonecode,ReColorCode"
            SQL = SQL + " from " + TABLE + " a,V_NewColor b"
            SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'  "
            SQL = SQL + wSts + wFormNo + wDivision + wNo + wName + wBuyer + wCustomer + wYKKColorType + wYKKColorcode + wSLDColor + wVFColor + wNewOldColor
            SQL = SQL + ASdate1 + AEdate1 + wCustomerColor + wCustomerColorCode + wOverSeaYkkcode + wPANTONECode + wReColorCode + FSdate1 + FEdate1
        Else
            Datagrid2.Columns.Item(0).HeaderText = "編號"
            Datagrid2.Columns.Item(1).HeaderText = "依賴狀態"
            Datagrid2.Columns.Item(2).HeaderText = "依賴日期"
            Datagrid2.Columns.Item(3).HeaderText = "完成日"

            Datagrid2.Columns.Item(4).HeaderText = "依賴表單"
            Datagrid2.Columns.Item(5).HeaderText = "依賴部門"
            Datagrid2.Columns.Item(6).HeaderText = "依賴者"

            Datagrid2.Columns.Item(7).HeaderText = "YKK色別"
            Datagrid2.Columns.Item(8).HeaderText = "YKK色號"
            Datagrid2.Columns.Item(9).HeaderText = "兼用色拉頭"
            Datagrid2.Columns.Item(10).HeaderText = "兼用色VF上/下止"
            Datagrid2.Columns.Item(11).HeaderText = "新舊色"



            SQL = "SELECT "
            SQL = SQL + " b.No  As Field1,"
            SQL = SQL + " case b.Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2, "
            SQL = SQL + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3, "
            SQL = SQL + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate, "

            SQL = SQL + " case when a.formno ='005008' and convert(char(10),getdate(),111) >= '2025/04/01' then '型別轉換-05CNPBS16' else  a.FormName  end as Field4,"
            SQL = SQL + " a.Divname As Field5,a.ApplyName as Field6,"
            SQL = SQL + " YKKColorType As Field7,YKKColorCode as Field8,SLDColor As Field9,VFColor As Field10,NewOldColor As Field11,"
            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            If DTABLE.SelectedValue = "未封存" Then
                SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
                SQL = SQL + "'pFormNo='   + a.FormNo + "
                SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
                SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
                SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
                SQL = SQL + "'&pApplyID=' + a.ApplyID "
                SQL = SQL + " As OPURL,  "
            Else
                SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
                SQL = SQL + "'pFormNo='   + a.FormNo + "
                SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
                SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
                SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
                SQL = SQL + "'&pApplyID=' + a.ApplyID  + '&pKeepdata=1'"
                SQL = SQL + " As OPURL,  "
            End If

            SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,DTSHEET"
            SQL = SQL + " from " + TABLE + " a, F_NewColorUAKIPLINGDT b"
            SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'  "
            SQL = SQL + wSts + wFormNo + wDivision + wNo + wName + wBuyer + wCustomer + wYKKColorType + wYKKColorcode + wSLDColor + wVFColor + wNewOldColor
            SQL = SQL + ASdate1 + AEdate1 + wCustomerColor + wCustomerColorCode + wOverSeaYkkcode + wPANTONECode + FSdate1 + FEdate1
        End If



        Dim SQL1 As String

        SQL1 = " SELECT * FROM (" + SQL + ")A WHERE 1=1 "
        SQL1 = SQL1 + AFSdate1 + AFEdate1
        SQL1 = SQL1 + " order by Field1 DESC  "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL1)
        If DFormNo2.Text <> "005012" Then
            DataGrid1.Visible = True
            Datagrid2.Visible = False
            DataGrid1.DataSource = DBAdapter1
            DataGrid1.DataBind()
            If DataGrid1.Items.Count > 0 Then
                BExcel.Visible = True
            Else
                BExcel.Visible = False

            End If

        Else
            Datagrid2.Visible = True
            DataGrid1.Visible = False
            Datagrid2.DataSource = DBAdapter1
            Datagrid2.DataBind()
            If Datagrid2.Items.Count > 0 Then
                BExcel.Visible = True
            Else
                BExcel.Visible = False

            End If
        End If




    End Sub

    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Protected Sub BClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClear.Click
        DSts1.Text = ""
        DNo.Text = ""
        DFormNo1.Text = ""
        DName.Text = ""
        DCustomer.Text = ""
        DBuyer.Text = ""
        DDivision1.Text = ""
        DYKKColortype.Text = ""
        DYKKColorCode.Text = ""
        DSLDColor.Text = ""
        DVFColor.Text = ""
        DASDate.Text = ""
        DAEDate.Text = ""
        DFSDate.Text = ""
        DFEDate.Text = ""
        DCustomerColor.Text = ""
        DCustomerColorCode.Text = ""
        DOverSeaYkkcode.Text = ""
        DPANTONECode.Text = ""

    End Sub

    Protected Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub

    Protected Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound



        e.Item.Cells(10).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(11).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(12).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(13).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(14).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(15).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        'e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        'e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")

    End Sub

    Protected Sub Datagrid2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Datagrid2.SelectedIndexChanged

    End Sub
End Class

