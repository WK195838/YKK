Imports System.Data
Imports System.Data.OleDb


Partial Class StockdtList
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
    Dim qty As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        wStep = Request.QueryString("pStep")
 
        If Not Me.IsPostBack Then   '不是PostBack
            If Request.QueryString("pKey") <> "" Then
                InputData = Request.QueryString("pKey")
                '     GetData()
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

        'InputData = DData.Text
        Dim Input As String
        Input = DData.Text


        If Input <> "" Then
            If Mid(Input, 1, 3) = "09C" Or Mid(Input, 1, 3) = "SLD" Or Mid(Input, 1, 3) = "SPL" Or Mid(Input, 1, 3) = "PN" Or Mid(Input, 1, 3) = "VF" Then
                InputData = " and  WSHCXA ='" + DData.Text + "'"
                SQL = "SELECT REM2XA,WSHCXA,REM1XA,ITMCXA,IT1IA0,IT2IA0,CLRCXA,SUM(WQTYXA)WQTYXA,QUNCXA,MAX(RCD1XA)RCD1XA " _
                & "FROM WAVEALIB.TWS830A LEFT JOIN WAVEDLIB.FA000 ON ITMCXA = ITMCA0 " _
                & "WHERE RCD2XA = '0' AND " _
                & "      REM2XA = 'C' " _
                & InputData _
                & "GROUP BY REM2XA,WSHCXA,REM1XA,ITMCXA,IT1IA0,IT2IA0,CLRCXA,QUNCXA " _
                & "ORDER BY REM2XA,WSHCXA,ITMCXA,CLRCXA "

            Else
                InputData = " and ITMCXA ='" + DData.Text + "'"
                SQL = "SELECT distinct '' as REM2XA,'' as WSHCXA,'' as REM1XA,ITMCXA,IT1IA0,IT2IA0,CLRCXA, '' as WQTYXA,'' as QUNCXA,'' as RCD1XA " _
                    & "FROM WAVEALIB.TWS830A LEFT JOIN WAVEDLIB.FA000 ON ITMCXA = ITMCA0 " _
                    & "WHERE RCD2XA = '0' AND " _
                    & "      REM2XA = 'C' " _
                    & InputData _
                    & "ORDER BY REM2XA,WSHCXA,ITMCXA,CLRCXA "
            End If
        End If





        OleDbConnection2.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "GetData")

        GridView1.Visible = True
        GridView1.DataSource = DBDataSet2
        GridView1.DataBind()



        OleDbConnection1.Close()
        OleDbConnection2.Close()
    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
      
        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
    
        Dim PALETENO As String = GridView1.SelectedRow.Cells(1).Text
        If PALETENO = "&nbsp;" Then
            PALETENO = ""
        End If
        Dim ITEMCODE As String = GridView1.SelectedRow.Cells(3).Text
        If ITEMCODE = "&nbsp;" Then
            ITEMCODE = ""
        End If
        Dim COLOR As String = GridView1.SelectedRow.Cells(6).Text
        If COLOR = "&nbsp;" Then
            COLOR = ""
        End If
        Dim QTY As String
        If GridView1.SelectedRow.Cells(7).Text <> "&nbsp;" Then
            QTY = CStr(CInt(GridView1.SelectedRow.Cells(7).Text))
        Else
            QTY = ""
        End If



        Dim js As String = ""
        '  js &= "var obj = window.opener.document.all.DTNo;"
        '  js &= "obj.value = '" & DTNo & "';"
        js &= "var obj = window.opener.document.all.DPALETENO;"
        js &= "obj.value = '" & PALETENO & "';"
        js &= "var obj = window.opener.document.all.DITEMCODE;"
        js &= "obj.value = '" & ITEMCODE & "';"
        ' js &= "var obj = window.opener.document.all.DPQTY;"
        ' js &= "obj.value = '" & PQTY & "';"
        js &= "var obj = window.opener.document.all.DQTY;"
        js &= "obj.value = '" & QTY & "';"
        js &= "var obj = window.opener.document.all.DCOLOR;"
        js &= "obj.value = '" & COLOR & "';"

        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)


    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

 
   
End Class
