Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Image
Imports System


Partial Class YKKColorShow
    Inherits System.Web.UI.Page
    Dim InputData As String
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim Color1, Color2 As String

 

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack

            DCOLOR1.Text = Request.QueryString("field")
            Color1 = String.Format("{0,5}", DCOLOR1.Text)  '補空白

            GetData()
        End If
    End Sub

    Sub GetData()

        Label1.Text = ""
        Color1 = String.Format("{0,5}", DCOLOR1.Text)  '補空白
        Color2 = String.Format("{0,5}", DCOLOR2.Text)  '補空白

        Dim R1, R2, G1, G2, B1, B2 As Integer
        Dim SQL, SQL1 As String
        '先找出布帶色號
        SQL = " Select * from F_DTMWRGBCOLOR WHERE COLORCODE ='" + Color1 + "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            DDarkLight1.Text = DBAdapter1.Rows(0).Item("DARKLIGHT")
            DR1.Text = DBAdapter1.Rows(0).Item("RCOLOR")
            DG1.Text = DBAdapter1.Rows(0).Item("GCOLOR")
            DB1.Text = DBAdapter1.Rows(0).Item("BCOLOR")
            R1 = DR1.Text
            G1 = DG1.Text
            B1 = DB1.Text
        Else

            R1 = 255
            G1 = 255
            B1 = 255
            Label1.Text = UCase(DCOLOR1.Text) + "不存在"
        End If

        'Dim UserID As String = Request.QueryString("userid")

        Panel1.BackColor = Color.FromArgb(R1, G1, B1)

        '再找拉頭色號
        If DData.Text <> "" Then
            SQL = " Select * from F_DTMWRGBCOLOR WHERE COLORTYPE ='塗裝' AND COLORCODE ='" + Color2 + "'"

            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter2.Rows.Count > 0 Then
                DDarkLight2.Text = DBAdapter2.Rows(0).Item("DARKLIGHT")
                DR2.Text = DBAdapter2.Rows(0).Item("RCOLOR")
                DG2.Text = DBAdapter2.Rows(0).Item("GCOLOR")
                DB2.Text = DBAdapter2.Rows(0).Item("BCOLOR")
                R2 = DR2.Text
                G2 = DG2.Text
                B2 = DB2.Text
            Else
                SQL1 = " Select * from F_DTMWRGBCOLOR WHERE COLORTYPE ='染色' AND COLORCODE ='" + Color2 + "'"
                Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL1)
                If DBAdapter3.Rows.Count > 0 Then
                    DDarkLight2.Text = DBAdapter3.Rows(0).Item("DARKLIGHT")
                    DR2.Text = DBAdapter3.Rows(0).Item("RCOLOR")
                    DG2.Text = DBAdapter3.Rows(0).Item("GCOLOR")
                    DB2.Text = DBAdapter3.Rows(0).Item("BCOLOR")
                    R2 = DR2.Text
                    G2 = DG2.Text
                    B2 = DB2.Text
                Else
                    DDarkLight2.Text = ""
                    DR2.Text = ""
                    DG2.Text = ""
                    DB2.Text = ""
                    R2 = 255
                    G2 = 255
                    B2 = 255
                    DCOLOR2.Text = ""
                    Label1.Text = UCase(DCOLOR2.Text) + "不存在"
                End If
            End If
        Else
            R2 = 255
            G2 = 255
            B2 = 255
            DR2.Text = ""
            DG2.Text = ""
            DB2.Text = ""

        End If


        Panel2.BackColor = Color.FromArgb(R2, G2, B2)



    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = DData.Text
        DCOLOR2.Text = DData.Text
        GetData()
    End Sub


    Protected Sub BCHECK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BCHECK.Click
        Dim js As String = ""
        js &= "var obj = window.opener.document.all.D3;"
        js &= "obj.value = '" & UCase(DCOLOR2.Text) & "';"
        js &= "var obj = window.opener.document.all.DAgain;"
        js &= "obj.value = '" & UCase(DDarkLight2.Text) & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)
    End Sub
End Class
