Imports System.Data
Imports System.Data.OleDb


Partial Class StockList
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
    Dim wqty As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        wStep = Request.QueryString("pStep")
        If wStep = 10 Then
            DData.Text = Request.QueryString("field")
            DColor.Text = Trim(Request.QueryString("field1"))
            wqty = Request.QueryString("field2")
        End If
        If DTNo = "" Then
            DTNo = Now.ToString("yyyyMMddHHmmss") '虛擬單號
        End If

        If Not Me.IsPostBack Then   '不是PostBack
            If Request.QueryString("pKey") <> "" Then
                InputData = Request.QueryString("pKey")
                GetData()
            End If
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        'Dim UserID As String = Request.QueryString("userid")

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDbConnection2 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("WAVESOLEDB")  'SQL連結設定
        OleDbConnection2.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("WAVESOLEDB")  'SQL連結設定

        InputData = DData.Text
        InputData1 = String.Format("{0,5}", DColor.Text)  '補空白
        SQL = " select ITMCXA,CLRCXA,QUNCXA,WQTYXA,RCD1XA,RCT1XA,WSHCXA,REM1XA"
        'SQL = SQL + " Case when left(WSHCXA,2) ='PO' then 'ITEM+數量' when  left(WSHCXA,2) ='PR' or  left(WSHCXA,3) in('DYE','BOX') OR  left(WSHCXA,4) ='OVER' then 'ITEM+板數' else '板號' end as Type "
        SQL = SQL + " from WAVEALIB.TWS830A "
        SQL = SQL + "where 1=1 "
        'RCD1XA=0 and REM2XA ='C' and RCD2XA=0  "


        'Com_Code in ('10','11','12') "
        'SQL = SQL + "and substring(Dep_Code,1,1) = '1' "
        ' SQL = SQL + "and leav_cd <> 'Y' "
        If InputData <> "" Then
            SQL = SQL + " and  ITMCXA ='" + InputData + "'"

        End If
        If Trim(InputData1) <> "" Then
            SQL = SQL + " and  CLRCXA ='" + InputData1 + "'"

        Else
            SQL = SQL + "  "
        End If
        SQL = SQL + " order by RCD1XA,RCT1XA,WSHCXA  "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "GetData")


        If DBDataSet1.Tables("GetData").Rows.Count > 0 Then
            If ((Mid(DBDataSet1.Tables("GetData").Rows(0).Item("WSHCXA"), 1, 2) = "PO") Or (Mid(DBDataSet1.Tables("GetData").Rows(0).Item("WSHCXA"), 1, 2) = "PR") Or _
            (Mid(DBDataSet1.Tables("GetData").Rows(0).Item("WSHCXA"), 1, 3) = "BOX") Or (Mid(DBDataSet1.Tables("GetData").Rows(0).Item("WSHCXA"), 1, 3) = "DYE") Or _
            (Mid(DBDataSet1.Tables("GetData").Rows(0).Item("WSHCXA"), 1, 4) = "OVER")) And wStep = "" Then
                SQL = " select distinct  ITMCXA,CLRCXA,Type from ("
                SQL = SQL + " select ITMCXA,CLRCXA,QUNCXA,WQTYXA,RCD1XA,RCT1XA,WSHCXA,REM1XA,"
                SQL = SQL + " Case when left(WSHCXA,2) ='PO' then 'ITEM+數量' when  left(WSHCXA,2) ='PR' or  left(WSHCXA,3) in('DYE','BOX') OR  left(WSHCXA,4) ='OVER' then 'ITEM+板數' else '板號' end as Type "
                SQL = SQL + " from WAVEALIB.TWS830A )a "
                SQL = SQL + "where 1=1 "
                'RCD1XA=0 and REM2XA ='C' and RCD2XA=0  "



                'Com_Code in ('10','11','12') "
                'SQL = SQL + "and substring(Dep_Code,1,1) = '1' "
                ' SQL = SQL + "and leav_cd <> 'Y' "
                If InputData <> "" Then
                    SQL = SQL + " and  ITMCXA ='" + InputData + "'"

                End If
                If Trim(InputData1) <> "" Then
                    SQL = SQL + " and  CLRCXA ='" + InputData1 + "'"

                Else
                    SQL = SQL + "  "
                End If


                OleDbConnection2.Open()
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet2, "GetData")

                GridView1.Visible = True
                GridView2.Visible = False
                GridView1.DataSource = DBDataSet2
                GridView1.DataBind()
                BAdd.Visible = False



            Else


                GridView2.Visible = True
                GridView1.Visible = False
                GridView2.DataSource = DBDataSet1
                GridView2.DataBind()
                BAdd.Visible = True
            End If
        End If



        OleDbConnection1.Close()
        OleDbConnection2.Close()
    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = DData.Text
        InputData1 = String.Format("{0,5}", DColor.Text)  '補空白
        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
        Dim ItemCode As String = GridView1.SelectedRow.Cells(1).Text
        Dim Color As String = GridView1.SelectedRow.Cells(2).Text

        If Color = "&nbsp;" Then
            Color = ""
        End If
        Dim wtype As String = GridView1.SelectedRow.Cells(3).Text
      

        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DType;"
        js &= "obj.value = '" & wType & "';"
        js &= "var obj = window.opener.document.all.DCODE;"
        js &= "obj.value = '" & ItemCode & "';"
        js &= "var obj = window.opener.document.all.DColor;"
        js &= "obj.value = '" & Trim(Color) & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)


    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
 
    Protected Sub BAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BAdd.Click
        '加入棧板號
        Dim sql As String

        '先清掉暫存資料
        sql = " delete from F_StockOutSheetdttemp"
        sql = sql + " where no ='" + DTNo + "'"
        uDataBase.ExecuteNonQuery(sql)


        Dim cun As Integer = 0
        Dim i As Integer
        Dim PaletNo As String = ""
        Dim ItemCode As String = ""
        Dim Color As String = ""
        Dim qty As String = ""
        For i = 0 To Me.GridView2.Rows.Count - 1 Step i + 1

            PaletNo = GridView2.DataKeys(i).Value.ToString()
        
            ItemCode = Me.GridView2.Rows(i).Cells(1).Text
            Color = Me.GridView2.Rows(i).Cells(2).Text
            qty = Me.GridView2.Rows(i).Cells(4).Text

            If Color = "&nbsp;" Then
                Color = ""
            End If


            If (CType(GridView2.Rows(i).FindControl("CheckNo"), CheckBox)).Checked Then

                sql = " Insert into F_StockOutSheetdttemp (No,type, StockNo,boxno,itemcode, Qty,checkstockout,wingsts,stockOutdate,Createuser,CreateTime,ModifyUser,ModifyTime) "
                sql = sql & "VALUES( "
                sql = sql & " '" & DTNo & "', "
                sql = sql & " '',"
                sql = sql & " '" & PaletNo & "', "
                sql = sql & " '', "
                sql = sql & " '" + ItemCode + "', "
                sql = sql & " '" + qty + "', "
                sql = sql & " '0', "
                sql = sql & " '0', "
                sql = sql & " '', "
                sql = sql & " '" & Request.QueryString("pUserID") & "', "
                sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "', "
                sql = sql & " '" & Request.QueryString("pUserID") & "', "
                sql = sql & " '" & Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "
                uDataBase.ExecuteNonQuery(sql)

                'If PaletNo = "" Then
                '    PaletNo = GridView2.DataKeys(i).Value.ToString()
                'Else
                '    PaletNo = PaletNo + "," + GridView2.DataKeys(i).Value.ToString()
                'End If
                cun = cun + 1
            End If

        Next

        If cun > wqty Then
            Label1.Text = "申請者需要板數=" + qty + ",選取的板數=" + Str(cun) + "已超過！"
        Else
            If wType = "" Then
                wType = "板號"
            End If


            Dim js As String = ""
            js &= "var obj = window.opener.document.all.DTNo;"
            js &= "obj.value = '" & DTNo & "';"
            js &= "var obj = window.opener.document.all.DType;"
            js &= "obj.value = '" & wType & "';"
            js &= "var obj = window.opener.document.all.DCODE;"
            js &= "obj.value = '" & ItemCode & "';"
            js &= "var obj = window.opener.document.all.DColor;"
            js &= "obj.value = '" & Trim(Color) & "';"
            js &= "window.opener.document.forms[0].submit();"
            js &= "window.close();"

            Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

        End If
       
        'Dim a As String = PaletNo
        'Dim a_result() As String = a.Split(","c)

        'Dim row As System.String
        'For Each row In a_result
        '    Console.Write(row + "\n")
        'Next

    End Sub

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
     

        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DPaletNo;"
        js &= "obj.value = '" & DTNo & "';"
     
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub

End Class
