Imports System.Data
Imports System.Data.OleDb


Partial Class ApproveNoList
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim InputData As String
    Dim InputData1 As String
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Dim DTNo As String = ""
    Dim wStep As String = ""
    Dim wType As String = ""
    Dim wqty As String = "0"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Sub GetData()


        Dim SQL As String

        SQL = " select  b.Unique_ID,b.no+'-'+seqno as no,supplier,code,size,family,body,puller,color,finish,case when Oldcancel ='1' then '廢除' else '' end as OldCancel  from  F_QASheet a, F_QASheetDT b "
        SQL = SQL + " where a.no =b.no  and sts in (0,1)    and  size++family+body+puller+color+finish  like '%" + InputData + "%'"
        SQL = SQL + " union all "
        SQL = SQL + "  select unique_id,acceptedno,partner,'' as code,size,family,body,puller,color,finish ,"
        SQL = SQL + "  case when  formsno =1 then '廢除' else '' end as OldCancel  from M_SPDIRWEDX"
        SQL = SQL + "  where STS=1 AND CAT='QC' and  size++family+body+puller+color+finish  like '%" + InputData + "%'"

        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()


    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = DData.Text
        '字串取代
        InputData = InputData.Replace(" ", "")
        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
        Dim ApproveNo As String = GridView1.SelectedRow.Cells(2).Text


        If GridView1.SelectedRow.Cells(11).Text = "廢除" Then
            uJavaScript.PopMsg(Me, "已廢除不可選取！")
        Else

            Dim js As String = ""

            js &= "var obj = window.opener.document.all.DApproveNo;"
            js &= "obj.value = '" & ApproveNo & "';"

            'If GridView1.SelectedRow.Cells(3).Text <> "&nbsp;" Then
            '    js &= "var obj = window.opener.document.all.DSupplier;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(3).Text & "';"
            'End If

            'If GridView1.SelectedRow.Cells(5).Text <> "&nbsp;" Then
            '    js &= "var obj = window.opener.document.all.DSize;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(5).Text & "';"
            'End If

            'If GridView1.SelectedRow.Cells(6).Text <> "&nbsp;" Then
            '    js &= "var obj = window.opener.document.all.DFamily;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(6).Text & "';"
            'End If

            'If GridView1.SelectedRow.Cells(7).Text <> "&nbsp;" Then
            '    js &= "var obj = window.opener.document.all.DBody;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(7).Text & "';"
            'End If

            'If GridView1.SelectedRow.Cells(8).Text <> "&nbsp;" Then
            '    ApproveNo = ApproveNo + "-" + GridView1.SelectedRow.Cells(8).Text
            '    js &= "var obj = window.opener.document.all.DPuller;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(8).Text & "';"
            'End If

            'If GridView1.SelectedRow.Cells(9).Text <> "&nbsp;" Then
            '    ApproveNo = ApproveNo + "-" + GridView1.SelectedRow.Cells(9).Text
            '    js &= "var obj = window.opener.document.all.DColor;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(9).Text & "';"
            'End If

            'If GridView1.SelectedRow.Cells(10).Text <> "&nbsp;" Then
            '    js &= "var obj = window.opener.document.all.DFinish;"
            '    js &= "obj.value = '" & GridView1.SelectedRow.Cells(10).Text & "';"
            'End If


            js &= "var obj = window.opener.document.all.DAppNO;"
            js &= "obj.value = '" & ApproveNo & "';"

            js &= "window.opener.document.forms[0].submit();"
            js &= "window.close();"


            Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)
        End If





    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(1).Visible = False
    End Sub
End Class
