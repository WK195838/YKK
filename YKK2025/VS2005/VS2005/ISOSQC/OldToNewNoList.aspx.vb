Imports System.Data
Imports System.Data.OleDb


Partial Class OldToNewNoList
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
 
        SQL = " select * from ( select  b.Unique_ID,b.no+'-'+seqno as no,supplier,code,size,family,body,puller,color,finish,"
        SQL = SQL + " case when Oldcancel ='1' then '廢除' else '' end as OldCancel,"
        SQL = SQL + " case when result1 ='OK' or result2 ='OK' then 'OK' else  ''  end as result1 "
        SQL = SQL + " from  F_QASheet a, F_QASheetDT b "

        SQL = SQL + " where a.no =b.no  and sts in (1,0)    and  size++family+body+puller+color+finish  like '%" + InputData + "%'   "
        SQL = SQL + " union all "
        SQL = SQL + "  select unique_id,acceptedno,partner,'' as code,size,family,body,puller,color,finish ,"
        SQL = SQL + "  case when  formsno =1 then '廢除' else '' end as OldCancel,Result1  from M_SPDIRWEDX"
        SQL = SQL + "  where STS=1 AND CAT='QC' and  size++family+body+puller+color+finish  like '%" + InputData + "%' )a order by no"


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

 
        If GridView1.SelectedRow.Cells(11).Text = "廢除" Then
            uJavaScript.PopMsg(Me, "已廢除不可選取！")
        ElseIf GridView1.SelectedRow.Cells(12).Text <> "OK" Then
            uJavaScript.PopMsg(Me, "判定OK才可選取！")

        Else
            Dim fpObj As New ForProject
            Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
            Dim OldToNewNo As String = GridView1.SelectedRow.Cells(2).Text


            Dim js As String = ""


            If GridView1.SelectedRow.Cells(3).Text <> "&nbsp;" Then
                js &= "var obj = window.opener.document.all.DSupplier;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(3).Text & "';"
            End If

            If GridView1.SelectedRow.Cells(5).Text <> "&nbsp;" Then
                js &= "var obj = window.opener.document.all.DSize;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(5).Text & "';"
            End If

            If GridView1.SelectedRow.Cells(6).Text <> "&nbsp;" Then
                js &= "var obj = window.opener.document.all.DFamily;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(6).Text & "';"
            End If

            If GridView1.SelectedRow.Cells(7).Text <> "&nbsp;" Then
                js &= "var obj = window.opener.document.all.DBody;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(7).Text & "';"
            End If

            If GridView1.SelectedRow.Cells(8).Text <> "&nbsp;" Then
                OldToNewNo = OldToNewNo + "-" + GridView1.SelectedRow.Cells(8).Text
                js &= "var obj = window.opener.document.all.DPuller;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(8).Text & "';"
            End If

            If GridView1.SelectedRow.Cells(9).Text <> "&nbsp;" Then
                OldToNewNo = OldToNewNo + "-" + GridView1.SelectedRow.Cells(9).Text
                js &= "var obj = window.opener.document.all.DColor;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(9).Text & "';"
            End If

            If GridView1.SelectedRow.Cells(10).Text <> "&nbsp;" Then
                js &= "var obj = window.opener.document.all.DFinish;"
                js &= "obj.value = '" & GridView1.SelectedRow.Cells(10).Text & "';"
            End If

            js &= "var obj = window.opener.document.all.DOldNO;"
            js &= "obj.value = '" & OldToNewNo & "';"



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
