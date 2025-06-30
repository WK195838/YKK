Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_COMBIinqCommission
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
        BSts.Attributes("onclick") = "GetStsPicker();"
        BFormNo.Attributes("onclick") = "GetCombiItemPicker();"
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

        ' pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=DTMW_COMBICommission_ist.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '程式別不同


      
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
            FSdate1 = " and completedDate >= '" + FSdate1 + "'"
        End If

        Dim FEdate As Date
        Dim FEdate1 As String = ""

        If DFEDate.Text <> "" Then
            FEdate = DFEDate.Text
            FEdate1 = Format(FEdate, "yyyy/MM/dd")
            FEdate1 = " and CompletedDate <= '" + FEdate1 + "'"
        End If

        Dim wSts As String = ""
        wSts = DSts2.Text
        If DSts1.Text <> "" Then
            wSts = " and b.Sts =" + wSts
        Else
            wSts = ""
        End If



        '取項目字串
        Dim wCombiItem As String = ""
        If DCombiItem1.Text <> "" And DCombiItem1.Text <> "全部" Then

            wCombiItem = DCombiItem1.Text
            Dim spiltResult As String() = wCombiItem.Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries)
            Dim cun As Integer
            cun = spiltResult.Length
            For cun = 0 To cun - 1
                If cun = 0 Then
                    wCombiItem = "'" + spiltResult(cun) + "'"
                Else
                    wCombiItem = wCombiItem + ",'" + spiltResult(cun) + "'"
                End If
            Next

            wCombiItem = " and CombiItem in (" + wCombiItem + ")"
        Else
            wCombiItem = ""
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
            wNo = " and  a.No like '%" + DNo.Text + "%'"
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
            wCustomer = " and b.Customer like '%" + DCustomer.Text + "%'"
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
 
        Dim wLTape As String = ""
        If DLTape.Text <> "" Then
            wLTape = " and Field12 like '%" + DLTape.Text + "%' "
        Else
            wLTape = ""
        End If


   

        Dim wLChain As String = ""
        If DLChain.Text <> "" Then
            wLChain = " and Field13 like '%" + DLChain.Text + "%' "
        Else
            wLChain = ""
        End If


        Dim wRChain As String = ""
        If DRChain.Text <> "" Then
            wRChain = " and Field14 like '%" + DRChain.Text + "%' "
        Else
            wRChain = ""
        End If

        Dim wRTape As String = ""
        If DRTape.Text <> "" Then
            wRTape = " and Field15 like '%" + DRTape.Text + "%' "
        Else
            wRTape = ""
        End If



        Dim i As Integer = 0
        Dim SQL As String

        DataGrid1.Columns.Item(0).HeaderText = "編號"
        DataGrid1.Columns.Item(1).HeaderText = "依賴狀態"
        DataGrid1.Columns.Item(2).HeaderText = "依賴日期"
        DataGrid1.Columns.Item(3).HeaderText = "完成日"

        DataGrid1.Columns.Item(4).HeaderText = "依賴表單"
        DataGrid1.Columns.Item(5).HeaderText = "依賴部門"
        DataGrid1.Columns.Item(6).HeaderText = "依賴者"
        DataGrid1.Columns.Item(7).HeaderText = "客戶"
        DataGrid1.Columns.Item(8).HeaderText = "BUYER"
        DataGrid1.Columns.Item(9).HeaderText = "YKK色別"
        DataGrid1.Columns.Item(10).HeaderText = "YKK色號"


        Dim TABLE As String = ""
        If DTABLE.SelectedValue = "未封存" Then
            TABLE = "V_WaitHandle_01"
        Else
            TABLE = "V_WaitHandle_old_01"
        End If



        SQL = "SELECT "
        SQL = SQL + " a.No  As Field1,"
        SQL = SQL + " case b.Sts When '0' Then '核定中' When '1' Then '完成' When '2' Then '取消' Else '取消' End As Field2, "
        SQL = SQL + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3, "
        SQL = SQL + "  case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate, "
        SQL = SQL + " a.FormName as Field4,"
        SQL = SQL + " a.Division As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"
        SQL = SQL + " YKKColorType As Field9,YKKColorCode as Field10,"
        SQL = SQL + " combiitem as Field11,"
        SQL = SQL + " case when vfltape  <> '' then  vfltape "
        SQL = SQL + " when vfmltape <> '' then vfmltape "
        SQL = SQL + " when pfmfltape <> '' then pfmfltape "
        SQL = SQL + " else '' end as Field12,"
        SQL = SQL + " case when vflchain  <> '' then  vflchain "
        SQL = SQL + " when vfmlchain <> '' then vfmlchain "
        SQL = SQL + " else '' end as Field13,"
        SQL = SQL + " case when vfrchain  <> '' then  vfrchain "
        SQL = SQL + " when vfmrchain <> '' then vfmrchain "
        SQL = SQL + " else '' end as Field14,"
        SQL = SQL + " case when vfrtape  <> '' then  vfrtape "
        SQL = SQL + " when vfmrtape <> '' then vfmrtape "
        SQL = SQL + " when pfmfrtape <> '' then pfmfrtape "
        SQL = SQL + " else '' end as Field15,"
        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + a.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(a.FormSno,Len(a.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(a.Step,Len(a.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + a.ApplyID "
        SQL = SQL + " As OPURL,  "
        SQL = SQL + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate"
        SQL = SQL + " from " + TABLE + " a,F_CombiSheet b"
        SQL = SQL + " Where a.formno=b.formno and a.formsno =b.formsno and Step  = '1'  "
        SQL = SQL + wSts + wCombiItem + wDivision + wNo + wName + wBuyer + wYKKColorType + wYKKColorcode + wCustomer
        SQL = SQL + ASdate1 + AEdate1
        Dim SQL1 As String

        SQL1 = " SELECT * FROM (" + SQL + ")A WHERE 1=1 "
        SQL1 = SQL1 + FSdate1 + FEdate1 + wLTape + wRTape + wLChain + wRChain
        SQL1 = SQL1 + " order by Field1 "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL1)
        DataGrid1.DataSource = DBAdapter1
        DataGrid1.DataBind()
        If DataGrid1.Items.Count > 0 Then
            BExcel.Visible = True
        Else
            BExcel.Visible = False

        End If

    End Sub

    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Protected Sub DSts1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSts1.TextChanged

    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Protected Sub BClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClear.Click
        DSts1.Text = ""
        DNo.Text = ""
        DCombiItem1.Text = ""
        DCombiItem2.Text = ""
        DDivision1.Text = ""
        DDivision2.Text = ""
        DName.Text = ""
        DCustomer.Text = ""
        DBuyer.Text = ""
        DYKKColortype.Text = ""
        DYKKColorCode.Text = ""
        DLTape.Text = ""
        DRTape.Text = ""
        DLChain.Text = ""
        DRChain.Text = ""
        DASDate.Text = ""
        DAEDate.Text = ""
        DFSDate.Text = ""
        DFEDate.Text = ""

    End Sub

    Protected Sub DCombiItem1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCombiItem1.TextChanged

    End Sub

    Protected Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub

    Protected Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        e.Item.Cells(10).Attributes.Add("style", "vnd.ms-excel.numberformat:@")


        e.Item.Cells(12).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(13).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(14).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(15).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Item.Cells(16).Attributes.Add("style", "vnd.ms-excel.numberformat:@")

    End Sub
End Class

