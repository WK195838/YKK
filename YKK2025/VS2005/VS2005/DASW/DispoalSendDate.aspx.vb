Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DispoalSendDate
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

        SetParameter()
        If Not IsPostBack Then

            GETSETDATE()

        End If

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


   

    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

        BPCE1.Attributes("onclick") = "calendarPicker('Form1.DPCE1');"
        BPCE2.Attributes("onclick") = "calendarPicker('Form1.DPCE2');"
        BPCE3.Attributes("onclick") = "calendarPicker('Form1.DPCE3');"


    End Sub



    Sub GETSETDATE()

        DPCS1.Text = ""
        DPCS2.Text = ""
        DPCS3.Text = ""

        DPCE1.Text = ""
        DPCE2.Text = ""
        DPCE3.Text = ""

        DAdmin1.Text = ""
        DAdmin2.Text = ""
        DAdmin3.Text = ""

        Dim DKEY As String = ""


        '寫入生管簽核截止日期
        Dim SQL As String
        SQL = "select  substring(dkey,10,len(dkey)-1)dkey,data,SUBSTRING(DATA,9,10) PCDATE,"
        SQL = SQL + "  substring(convert(char(10),getdate(),111),1,7) YM1 ,"
        SQL = SQL + "  substring(convert(char(10),dateadd(month,1,getdate()),111),1,7)YM2,"
        SQL = SQL + "  substring(convert(char(10),dateadd(month,2,getdate()),111),1,7) YM3"
        SQL = SQL + "  from  m_referp"
        SQL = SQL + " where  cat = 6001"
        SQL = SQL + " AND LEFT(DATA,7) >=  substring(convert(char(10),getdate(),111),1,7)  "
        '  SQL = SQL + " AND LEFT(DATA,7) >='2015/11' "
        SQL = SQL + " and left(dkey,12) ='SendTime-PCS' ORDER BY DATA "
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        Dim I As Integer
        Dim SQL1 As String

        If dtFlow.Rows.Count > 0 Then
            LYM1.Text = dtFlow.Rows(0)("YM1")
            LYM2.Text = dtFlow.Rows(0)("YM2")
            LYM3.Text = dtFlow.Rows(0)("YM3")

            For I = 0 To dtFlow.Rows.Count - 1

                If I = 0 Then
                    DPCS1.Text = dtFlow.Rows(I).Item("PCDATE")
                ElseIf I = 1 Then
                    DPCS2.Text = CStr(dtFlow.Rows(I).Item("PCDATE"))
                Else
                    DPCS3.Text = dtFlow.Rows(I).Item("PCDATE")
                End If

                DKEY = Mid(dtFlow.Rows(I).Item("DKEY"), 4, 1)


                '寫入生管簽核申請日期
                SQL1 = "select  substring(dkey,10,len(dkey)-1)dkey,data,SUBSTRING(DATA,9,10) PCDATE,"
                SQL1 = SQL1 + "  substring(convert(char(10),getdate(),111),1,7) YM1 ,"
                SQL1 = SQL1 + "  substring(convert(char(10),dateadd(month,1,getdate()),111),1,7)YM2,"
                SQL1 = SQL1 + "  substring(convert(char(10),dateadd(month,2,getdate()),111),1,7) YM3"
                SQL1 = SQL1 + "  from  m_referp"
                SQL1 = SQL1 + " where  cat = 6001"
                SQL1 = SQL1 + " AND LEFT(DATA,7) >=  substring(convert(char(10),getdate(),111),1,7)  "
                SQL1 = SQL1 + " and left(dkey,13) ='SendTime-PCE" + CStr(DKEY) + "'"
                Dim dtFlow3 As DataTable = uDataBase.GetDataTable(SQL1)


                If I = 0 Then

                    DPCE1.Text = dtFlow3.Rows(0).Item("DATA")
                ElseIf I = 1 Then

                    DPCE2.Text = dtFlow3.Rows(0).Item("DATA")
                ElseIf I = 2 Then

                    DPCE3.Text = dtFlow3.Rows(0).Item("DATA")
                End If





            Next

        End If

    
        '寫入廠長簽核申請日期
        SQL = "select  substring(dkey,10,len(dkey)-1)dkey,data,SUBSTRING(DATA,9,10) PCDATE,"
        SQL = SQL + "  substring(convert(char(10),getdate(),111),1,7) YM1 ,"
        SQL = SQL + "  substring(convert(char(10),dateadd(month,1,getdate()),111),1,7)YM2,"
        SQL = SQL + "  substring(convert(char(10),dateadd(month,2,getdate()),111),1,7) YM3"
        SQL = SQL + "  from  m_referp"
        SQL = SQL + " where  cat = 6001"
        SQL = SQL + " AND LEFT(DATA,7) >=  substring(convert(char(10),getdate(),111),1,7)  "
        ' SQL = SQL + " AND LEFT(DATA,7) >='2015/11' "
        SQL = SQL + " and left(dkey,14) ='SendTime-ADMIN' ORDER BY DATA "
        Dim dtFlow2 As DataTable = uDataBase.GetDataTable(SQL)

        If dtFlow2.Rows.Count > 0 Then

            For I = 0 To dtFlow2.Rows.Count - 1

                If I = 0 Then
                    DAdmin1.Text = dtFlow2.Rows(I).Item("PCDATE")
                ElseIf I = 1 Then
                    DAdmin2.Text = dtFlow2.Rows(I).Item("PCDATE")
                ElseIf I = 2 Then
                    DAdmin3.Text = dtFlow2.Rows(I).Item("PCDATE")
                End If

            Next

        End If


        '寫入立會日期
        SQL = "select  substring(dkey,10,len(dkey)-1)dkey,substring(data,1,7)data,SUBSTRING(DATA,9,10) ACCDATE,"
        SQL = SQL + "  substring(convert(char(10),getdate(),111),1,7) YM1 ,"
        SQL = SQL + "  substring(convert(char(10),dateadd(month,1,getdate()),111),1,7)YM2,"
        SQL = SQL + "  substring(convert(char(10),dateadd(month,2,getdate()),111),1,7) YM3"
        SQL = SQL + "  from  m_referp"
        SQL = SQL + " where  cat = 6001"
        SQL = SQL + " AND LEFT(DATA,7) >=  substring(convert(char(10),getdate(),111),1,7)  "
        ' SQL = SQL + " AND LEFT(DATA,7) >='2015/11' "
        SQL = SQL + " and left(dkey,11) ='AccountTime' ORDER BY DATA "
        Dim dtFlow4 As DataTable = uDataBase.GetDataTable(SQL)

        If dtFlow4.Rows.Count > 0 Then

            For I = 0 To dtFlow4.Rows.Count - 1

                If dtFlow4.Rows(I).Item("Data") = dtFlow4.Rows(I).Item("YM1") Then
                    DAccount1.Text = dtFlow4.Rows(I).Item("ACCDATE")
                ElseIf dtFlow4.Rows(I).Item("Data") = dtFlow4.Rows(I).Item("YM2") Then
                    DAccount2.Text = dtFlow4.Rows(I).Item("ACCDATE")
                ElseIf dtFlow4.Rows(I).Item("Data") = dtFlow4.Rows(I).Item("YM3") Then
                    DAccount3.Text = dtFlow4.Rows(I).Item("ACCDATE")
                End If

            Next

        End If



    End Sub
    Protected Sub BSetDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSetDate.Click
        If OK() = True Then
            Dim SQL As String
            Dim PCS1, PCS2, PCS3, ADMIN1, ADMIN2, ADMIN3, Account1, Account2, Account3 As String

            PCS1 = DPCS1.SelectedValue
            PCS2 = DPCS2.SelectedValue
            PCS3 = DPCS3.SelectedValue
            ADMIN1 = DAdmin1.SelectedValue
            ADMIN2 = DAdmin2.SelectedValue
            ADMIN3 = DAdmin3.SelectedValue

            Account1 = DAccount1.SelectedValue
            Account2 = DAccount2.SelectedValue
            Account3 = DAccount3.SelectedValue


            '生管截止日
            SQL = "update  m_referp"
            SQL = SQL + " set Data = '" + LYM1.Text + "/" + PCS1.PadLeft(2, "00") + "'"
            SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
            SQL = SQL + " ,modifyTime=getdate() "
            SQL = SQL + " where  cat = 6001"
            SQL = SQL + " and dkey ='SendTime-PCS1'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = "update  m_referp"
            SQL = SQL + " set Data =  '" + LYM2.Text + "/" + PCS2.PadLeft(2, "00") + "'"
            SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
            SQL = SQL + " ,modifyTime=getdate() "
            SQL = SQL + " where  cat = 6001"
            SQL = SQL + " and dkey ='SendTime-PCS2'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = "update  m_referp"
            SQL = SQL + " set Data =  '" + LYM3.Text + "/" + PCS3.PadLeft(2, "00") + "'"
            SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
            SQL = SQL + " ,modifyTime=getdate() "
            SQL = SQL + " where  cat = 6001"
            SQL = SQL + " and dkey ='SendTime-PCS3'"
            uDataBase.ExecuteNonQuery(SQL)


            '生管申請日
            If DPCE1.Text <> "" Then
                SQL = "update  m_referp"
                SQL = SQL + " set Data = '" + CDate(DPCE1.Text).ToString("yyyy/MM/dd") + "'"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='SendTime-PCE1'"
                uDataBase.ExecuteNonQuery(SQL)
            End If

            If DPCE2.Text <> "" Then
                SQL = "update  m_referp"
                SQL = SQL + " set Data =  '" + CDate(DPCE2.Text).ToString("yyyy/MM/dd") + "'"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='SendTime-PCE2'"
                uDataBase.ExecuteNonQuery(SQL)

            End If

            If DPCE3.Text <> "" Then
                SQL = "update  m_referp"
                SQL = SQL + " set Data = '" + CDate(DPCE3.Text).ToString("yyyy/MM/dd") + "'"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='SendTime-PCE3'"
                uDataBase.ExecuteNonQuery(SQL)

            End If


            '廠長簽核日期
            SQL = "update  m_referp"
            SQL = SQL + " set Data = '" + LYM1.Text + "/" + ADMIN1.PadLeft(2, "00") + "'"
            SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
            SQL = SQL + " ,modifyTime=getdate() "
            SQL = SQL + " where  cat = 6001"
            SQL = SQL + " and dkey ='SendTime-Admin1'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = "update  m_referp"
            SQL = SQL + " set Data ='" + LYM2.Text + "/" + ADMIN2.PadLeft(2, "00") + "'"
            SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
            SQL = SQL + " ,modifyTime=getdate() "
            SQL = SQL + " where  cat = 6001"
            SQL = SQL + " and dkey ='SendTime-Admin2'"
            uDataBase.ExecuteNonQuery(SQL)

            SQL = "update  m_referp"
            SQL = SQL + " set Data = '" + LYM3.Text + "/" + ADMIN3.PadLeft(2, "00") + "'"
            SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
            SQL = SQL + " ,modifyTime=getdate() "
            SQL = SQL + " where  cat = 6001"
            SQL = SQL + " and dkey ='SendTime-Admin3'"
            uDataBase.ExecuteNonQuery(SQL)

            '立會日期
            If DAccount1.SelectedValue <> "" Then
                SQL = "update  m_referp"
                SQL = SQL + " set Data = '" + LYM1.Text + "/" + Account1.PadLeft(2, "00") + "'"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='AccountTime1'"
                uDataBase.ExecuteNonQuery(SQL)
            Else
                SQL = "update  m_referp"
                SQL = SQL + " set Data = ''"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='AccountTime1'"
                uDataBase.ExecuteNonQuery(SQL)
            End If

            If DAccount2.SelectedValue <> "" Then
                SQL = "update  m_referp"
                SQL = SQL + " set Data ='" + LYM2.Text + "/" + Account2.PadLeft(2, "00") + "'"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='AccountTime2'"
                uDataBase.ExecuteNonQuery(SQL)
            Else
                SQL = "update  m_referp"
                SQL = SQL + " set Data =''"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='AccountTime2'"
                uDataBase.ExecuteNonQuery(SQL)
            End If

            If DAccount3.SelectedValue <> "" Then
                SQL = "update  m_referp"
                SQL = SQL + " set Data = '" + LYM3.Text + "/" + Account3.PadLeft(2, "00") + "'"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='AccountTime3'"
                uDataBase.ExecuteNonQuery(SQL)
            Else
                SQL = "update  m_referp"
                SQL = SQL + " set Data = ''"
                SQL = SQL + " ,modifyUser = '" + Request.Cookies("UserID").Value + "'"
                SQL = SQL + " ,modifyTime=getdate() "
                SQL = SQL + " where  cat = 6001"
                SQL = SQL + " and dkey ='AccountTime3'"
                uDataBase.ExecuteNonQuery(SQL)
            End If




            uJavaScript.PopMsg(Me, "報廢批次簽核日期設定完成")



        End If

       
    End Sub

    Function OK() As Boolean

        Dim PCS1, PCS2, PCS3, ADMIN1, ADMIN2, ADMIN3 As String
        PCS1 = DPCS1.SelectedValue
        PCS2 = DPCS2.SelectedValue
        PCS3 = DPCS3.SelectedValue
        ADMIN1 = DAdmin1.SelectedValue
        ADMIN2 = DAdmin2.SelectedValue
        ADMIN3 = DAdmin3.SelectedValue


        Dim isOK As Boolean = True
        Dim Message As String = ""


        If DPCS1.Text = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM1.Text + "請輸入截止日!"
        End If

        If DPCS2.Text = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM2.Text + "請輸入截止日!"
        End If


        If DPCS3.Text = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM3.Text + "請輸入截止日!"
        End If

        If DPCE1.Text = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM1.Text + "請輸入開放日!"
        End If

        If DPCE2.Text = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM2.Text + "請輸入開放日!"
        End If

        If DPCE3.Text = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM3.Text + "請輸入開放日!"
        End If


        If DAdmin1.SelectedValue = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM1.Text + "請輸入廠長簽核日!"
        End If


        If DAdmin2.SelectedValue = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM2.Text + "請輸入廠長簽核日!"
        End If

        If DAdmin3.SelectedValue = "" Then
            isOK = False
            Message = Message + "\n" + "異常：" + LYM3.Text + "請輸入廠長簽核日!"
        End If

        If DPCE1.Text <> "" Then
            Dim DPCE1Date As DateTime
            DPCE1Date = DPCE1.Text

            If DPCE1Date.ToString("yyyy/mm/dd") > LYM1.Text + "/" + PCS1.PadLeft(2, "00") Then
                isOK = False
                Message = Message + "\n" + "異常：" + LYM1.Text + "開放日必須大於截止日!"
            End If
        End If


        If DPCE2.Text <> "" Then
            Dim DPCE2Date As DateTime
            DPCE2Date = DPCE2.Text

            If DPCE2Date.ToString("yyyy/mm/dd") > LYM2.Text + "/" + PCS2.PadLeft(2, "00") Then
                isOK = False
                Message = Message + "\n" + "異常：" + LYM2.Text + "開放日必須大於截止日!"
            End If

        End If



        If DPCE3.Text <> "" Then
            Dim DPCE3Date As DateTime
            DPCE3Date = DPCE3.Text
            If DPCE3Date.ToString("yyyy/mm/dd") > LYM3.Text + "/" + PCS3.PadLeft(2, "00") Then
                isOK = False
                Message = Message + "\n" + "異常：" + LYM3.Text + "開放日必須大於截止日!"
            End If
        End If



        If DPCE1.Text <> "" Or DAdmin1.SelectedValue <> "" Or DPCS1.SelectedValue <> "" Then
            If LYM1.Text + "/" + ADMIN1.PadLeft(2, "00") <= LYM1.Text + "/" + PCS1.PadLeft(2, "00") Then
                isOK = False
                Message = Message + "\n" + "異常：" + LYM1.Text + "廠長簽核日必須大於截止日!"
            End If
        End If

        If DPCE2.Text <> "" Or DAdmin2.SelectedValue <> "" Or DPCS2.SelectedValue <> "" Then
            If LYM2.Text + "/" + ADMIN2.PadLeft(2, "00") <= LYM2.Text + "/" + PCS2.PadLeft(2, "00") Then
                isOK = False
                Message = Message + "\n" + "異常：" + LYM2.Text + "廠長簽核日必須大於截止日!"
            End If
        End If


        If DPCE3.Text <> "" Or DAdmin3.SelectedValue <> "" Or DPCS3.SelectedValue <> "" Then
            If LYM3.Text + "/" + ADMIN3.PadLeft(2, "00") <= LYM3.Text + "/" + PCS3.PadLeft(2, "00") Then
                isOK = False
                Message = Message + "\n" + "異常：" + LYM3.Text + "廠長簽核日必須大於截止日!"
            End If
        End If


        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If



        Return isOK


    End Function

End Class

