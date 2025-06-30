Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class ItemRegisterSheet_02
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代號
    Dim wUserID As String           '簽核者
    Dim NowDateTime As String       '現在日期時間

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Dim uJavaScript As New Utility.JScript

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            LWaitHandle.Visible = False
            If wStep = 40 Then Waithandle()
            ShowFormData()          ' 顯示表單資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代號
        wUserID = Request.QueryString("pUserID")    '簽核者
        '
        Response.Cookies("PGM").Value = "ItemRegisterSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(WaitHandle)
    '**     待處理有無
    '**
    '*****************************************************************
    Sub Waithandle()
        Dim SQL As String
        SQL = "Select URL From V_WaitHandle_01 "
        SQL &= "Where FormNo =  '" & wFormNo & "'"
        SQL &= "  And FormSno =  '" & CStr(wFormSno) & "'"
        SQL &= "  And Active = '1' "
        SQL &= "  And (Sts = '0'  Or  Sts = '4') "
        SQL &= "  And Step = '10' "
        SQL &= "  And DecideID = '" + wUserID + "' "
        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
        If dt_WaitHandle.Rows.Count > 0 Then
            LWaitHandle.NavigateUrl = dt_WaitHandle.Rows(0).Item("URL")
            LWaitHandle.Visible = True
        End If
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("ItemRegisterFilePath")
        Dim PathOld1 As String = uCommon.GetAppSetting("HttpOld")
        Dim PathOld2 As String = uCommon.GetAppSetting("ItemRegisterFilePath")
        Dim str, xKey As String
        Dim iSlider As String()
        Dim RtnCode As Integer = 0

        Dim SQL As String
        SQL = "Select * From F_ItemRegisterSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtItemRegisterSheet.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtItemRegisterSheet.Rows(0).Item("No")                           ' No
            DDate.Text = dtItemRegisterSheet.Rows(0).Item("Date")                       ' 申請日
            DName.Text = dtItemRegisterSheet.Rows(0).Item("Name")                       ' 申請人姓名
            DJobTitle.Text = dtItemRegisterSheet.Rows(0).Item("JobTitle")               ' 職稱
            DDivision.Text = dtItemRegisterSheet.Rows(0).Item("Division")               ' 部門
            DRCode.Text = dtItemRegisterSheet.Rows(0).Item("RCode")                     ' R Code
            DRItemName1.Text = dtItemRegisterSheet.Rows(0).Item("RItemName1")           ' R Item Name-1
            DRItemName2.Text = dtItemRegisterSheet.Rows(0).Item("RItemName2")           ' R Item Name-2
            DRItemName3.Text = dtItemRegisterSheet.Rows(0).Item("RItemName3")           ' R Item Name-3
            DRSize.Text = dtItemRegisterSheet.Rows(0).Item("RSize")                     ' R Size
            DRChain.Text = dtItemRegisterSheet.Rows(0).Item("RChain")                   ' R Chain
            DRClass.Text = dtItemRegisterSheet.Rows(0).Item("RClass")                   ' R Class
            DRTape.Text = dtItemRegisterSheet.Rows(0).Item("RTape")                     ' R Tape
            DRSlider1.Text = dtItemRegisterSheet.Rows(0).Item("RSlider1")               ' R Slider1
            DRFinish1.Text = dtItemRegisterSheet.Rows(0).Item("RFinish1")               ' R Finish1
            DRSlider2.Text = dtItemRegisterSheet.Rows(0).Item("RSlider2")               ' R Slider2
            DRFinish2.Text = dtItemRegisterSheet.Rows(0).Item("RFinish2")               ' R Finish2
            DRSRequest1.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest1")           ' R 特殊要求1
            DRSRequest2.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest2")           ' R 特殊要求2
            DRSRequest3.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest3")           ' R 特殊要求3
            DRSRequest4.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest4")           ' R 特殊要求4
            DRSRequest5.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest5")           ' R 特殊要求5
            DRSRequest6.Text = dtItemRegisterSheet.Rows(0).Item("RSRequest6")           ' R 特殊要求6
            DRFamily.Text = dtItemRegisterSheet.Rows(0).Item("RFamily")                 ' R Family Code
            DRST1.Text = dtItemRegisterSheet.Rows(0).Item("RST1")                       ' R 統計區分1
            DRST2.Text = dtItemRegisterSheet.Rows(0).Item("RST2")                       ' R 統計區分2
            DRST3.Text = dtItemRegisterSheet.Rows(0).Item("RST3")                       ' R 統計區分3
            DRST4.Text = dtItemRegisterSheet.Rows(0).Item("RST4")                       ' R 統計區分4
            DRST5.Text = dtItemRegisterSheet.Rows(0).Item("RST5")                       ' R 統計區分5
            DRST6.Text = dtItemRegisterSheet.Rows(0).Item("RST6")                       ' R 統計區分6
            DRST7.Text = dtItemRegisterSheet.Rows(0).Item("RST7")                       ' R 統計區分7
            DRNoDisplay.Text = dtItemRegisterSheet.Rows(0).Item("RNoDisplay")           ' R No Display
            ' PriceVersion
            DA001.Checked = False
            DA206.Checked = False
            DA211.Checked = False
            DA999.Checked = False
            DK206.Checked = False
            DK211.Checked = False
            If dtItemRegisterSheet.Rows(0).Item("A001") = 1 Then DA001.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("A206") = 1 Then DA206.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("A211") = 1 Then DA211.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("A999") = 1 Then DA999.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("K206") = 1 Then DK206.Checked = True
            If dtItemRegisterSheet.Rows(0).Item("K211") = 1 Then DK211.Checked = True
            DCode.Text = dtItemRegisterSheet.Rows(0).Item("Code")                       ' Code
            DItemName1.Text = dtItemRegisterSheet.Rows(0).Item("ItemName1")             ' Item Name-1
            DItemName2.Text = dtItemRegisterSheet.Rows(0).Item("ItemName2")             ' Item Name-2
            DItemName3.Text = dtItemRegisterSheet.Rows(0).Item("ItemName3")             ' Item Name-3
            DSize.Text = dtItemRegisterSheet.Rows(0).Item("Size")                       ' Size
            DChain.Text = dtItemRegisterSheet.Rows(0).Item("Chain")                     ' Chain
            DClass.Text = dtItemRegisterSheet.Rows(0).Item("Class")                     ' Class
            DTape.Text = dtItemRegisterSheet.Rows(0).Item("Tape")                       ' Tape
            DSlider1.Text = dtItemRegisterSheet.Rows(0).Item("Slider1")                 ' Slider1
            DFinish1.Text = dtItemRegisterSheet.Rows(0).Item("Finish1")                 ' Finish1
            DSlider2.Text = dtItemRegisterSheet.Rows(0).Item("Slider2")                 ' Slider2
            DFinish2.Text = dtItemRegisterSheet.Rows(0).Item("Finish2")                 ' Finish2
            DSRequest1.Text = dtItemRegisterSheet.Rows(0).Item("SRequest1")             ' 特殊要求1
            DSRequest2.Text = dtItemRegisterSheet.Rows(0).Item("SRequest2")             ' 特殊要求2
            DSRequest3.Text = dtItemRegisterSheet.Rows(0).Item("SRequest3")             ' 特殊要求3
            DSRequest4.Text = dtItemRegisterSheet.Rows(0).Item("SRequest4")             ' 特殊要求4
            DSRequest5.Text = dtItemRegisterSheet.Rows(0).Item("SRequest5")             ' 特殊要求5
            DSRequest6.Text = dtItemRegisterSheet.Rows(0).Item("SRequest6")             ' 特殊要求6
            DFamily.Text = dtItemRegisterSheet.Rows(0).Item("Family")                   ' Family Code
            DST1.Text = dtItemRegisterSheet.Rows(0).Item("ST1")                         ' 統計區分1
            DST2.Text = dtItemRegisterSheet.Rows(0).Item("ST2")                         ' 統計區分2
            DST3.Text = dtItemRegisterSheet.Rows(0).Item("ST3")                         ' 統計區分3
            DST4.Text = dtItemRegisterSheet.Rows(0).Item("ST4")                         ' 統計區分4
            DST5.Text = dtItemRegisterSheet.Rows(0).Item("ST5")                         ' 統計區分5
            DST6.Text = dtItemRegisterSheet.Rows(0).Item("ST6")                         ' 統計區分6
            DST7.Text = dtItemRegisterSheet.Rows(0).Item("ST7")                         ' 統計區分7
            DNoDisplay.Text = dtItemRegisterSheet.Rows(0).Item("NoDisplay")             ' No Display
            If dtItemRegisterSheet.Rows(0).Item("PriceApply") = 0 Then                  ' Price Descriptioon / PriceNo
                DPriceDescr.Text = ""
                DPriceNo.Text = ""
            End If
            If dtItemRegisterSheet.Rows(0).Item("PriceApply") = 1 Then                  ' Price Descriptioon / PriceNo
                DPriceDescr.Text = "非賣品不設定/進口成品/手動計算"
                DPriceNo.Text = ""
            End If
            If dtItemRegisterSheet.Rows(0).Item("PriceApply") = 2 Then                  ' Price Descriptioon / PriceNo
                DPriceDescr.Text = "已設定"
                DPriceNo.Text = dtItemRegisterSheet.Rows(0).Item("PriceNo")
            End If

            If Not Request.QueryString("pUserID") Is Nothing Then                            ' SLD Price
                If Request.QueryString("pUserID").ToUpper = "MK028" Or _
                   Request.QueryString("pUserID").ToUpper = "IT003" Then
                    DSLDPrice.Text = dtItemRegisterSheet.Rows(0).Item("SLDPrice")
                End If
            Else
                DSLDPrice.Text = ""
            End If

            DRemark.Text = dtItemRegisterSheet.Rows(0).Item("Remark")                   ' Remark

            If dtItemRegisterSheet.Rows(0).Item("Attachfile1") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtItemRegisterSheet.Rows(0).Item("Attachfile1"))
                LAttachfile1.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile1")   '附件
                LAttachfile1.Visible = True
            Else
                LAttachfile1.Visible = False
            End If

            If dtItemRegisterSheet.Rows(0).Item("Attachfile2") <> "" Then
                '封存檔還原  20213/16  jessica
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, dtItemRegisterSheet.Rows(0).Item("Attachfile2"))
                LAttachfile2.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile2")   '附件
                LAttachfile2.Visible = True
            Else
                LAttachfile2.Visible = False
            End If


            DBuyer.Text = dtItemRegisterSheet.Rows(0).Item("Buyer")                         ' BUYER
            DBuyerCode.Text = dtItemRegisterSheet.Rows(0).Item("BuyerCode")                         ' BUYER
            DCustomer.Text = dtItemRegisterSheet.Rows(0).Item("Customer")                         ' Customer
            DCustomerCode.Text = dtItemRegisterSheet.Rows(0).Item("CustomerCode")                         ' Customer
            DForUse.Text = dtItemRegisterSheet.Rows(0).Item("ForUse")
            DOrderDate.Text = dtItemRegisterSheet.Rows(0).Item("OrderDate")
            LOrderDate.Style("left") = 120 & "px"
            DOrderDate.Style("left") = 200 & "px"



            'ZIP申請有無
            If dtItemRegisterSheet.Rows(0).Item("ZIPApply") = 1 Then
                SQL = "Select * From F_ItemRegisterZIPSheet "
                SQL &= " Where FormNo =  '" & dtItemRegisterSheet.Rows(0).Item("ZIPFormNo").ToString.Trim & "'"
                SQL &= "   And FormSno =  '" & CStr(CInt(Mid(dtItemRegisterSheet.Rows(0).Item("ZIPNo"), 5))) & "'"
                SQL &= "   And Sts = '1' "
                Dim dtZIP As DataTable = uDataBase.GetDataTable(SQL)
                If dtZIP.Rows.Count > 0 Then
                    LZIPSheet.NavigateUrl = "ItemRegisterZIPSheet_02.aspx?pFormNo=" & dtItemRegisterSheet.Rows(0).Item("ZIPFormNo").ToString.Trim & _
                                                                     "&pFormSno=" & CStr(CInt(Mid(dtItemRegisterSheet.Rows(0).Item("ZIPNo"), 5)))
                    LZIPSheet.Visible = True
                Else
                    LZIPSheet.Visible = False
                End If
            Else
                LZIPSheet.Visible = False
            End If
            'SLD申請有無
            If dtItemRegisterSheet.Rows(0).Item("SLDApply") = 1 Then
                SQL = "Select * From F_ItemRegisterSLDSheet "
                SQL &= " Where FormNo =  '" & dtItemRegisterSheet.Rows(0).Item("SLDFormNo").ToString.Trim & "'"
                SQL &= "   And FormSno =  '" & CStr(CInt(Mid(dtItemRegisterSheet.Rows(0).Item("SLDNo"), 5))) & "'"
                SQL &= "   And Sts = '1' "
                Dim dtSLD As DataTable = uDataBase.GetDataTable(SQL)
                If dtSLD.Rows.Count > 0 Then
                    LSLDSheet.NavigateUrl = "ItemRegisterSLDSheet_02.aspx?pFormNo=" & dtItemRegisterSheet.Rows(0).Item("SLDFormNo").ToString.Trim & _
                                                                     "&pFormSno=" & CStr(CInt(Mid(dtItemRegisterSheet.Rows(0).Item("SLDNo"), 5)))
                    LSLDSheet.Visible = True
                Else
                    LSLDSheet.Visible = False
                End If
            Else
                LSLDSheet.Visible = False
            End If
            'CH申請有無
            If dtItemRegisterSheet.Rows(0).Item("CHApply") = 1 Then
                SQL = "Select * From F_ItemRegisterCHSheet "
                SQL &= " Where FormNo =  '" & dtItemRegisterSheet.Rows(0).Item("CHFormNo").ToString.Trim & "'"
                SQL &= "   And FormSno =  '" & CStr(CInt(Mid(dtItemRegisterSheet.Rows(0).Item("CHNo"), 5))) & "'"
                SQL &= "   And Sts = '1' "
                Dim dtCH As DataTable = uDataBase.GetDataTable(SQL)
                If dtCH.Rows.Count > 0 Then
                    LCHSheet.NavigateUrl = "ItemRegisterCHSheet_02.aspx?pFormNo=" & dtItemRegisterSheet.Rows(0).Item("CHFormNo").ToString.Trim & _
                                                                     "&pFormSno=" & CStr(CInt(Mid(dtItemRegisterSheet.Rows(0).Item("CHNo"), 5)))
                    LCHSheet.Visible = True
                Else
                    LCHSheet.Visible = False
                End If
            Else
                LCHSheet.Visible = False
            End If

            Dim xSize, xFamily As String
            Dim xStr As String = DSize.Text & "-" & DFamily.Text
            xStr = Replace(xStr, "05-CI", "05-CN")
            xStr = Replace(xStr, "05-CZ", "05-CN")
            xStr = Replace(xStr, "05-CS", "05-CN")
            xStr = Replace(xStr, "05-Y", "05-M")
            xStr = Replace(xStr, "03-CZ", "03-C")
            xStr = Replace(xStr, "03-CS", "03-C")
            '
            xSize = Mid(xStr, 1, InStr(xStr, "-") - 1)
            xFamily = Mid(xStr, InStr(xStr, "-") + 1, 9)
            '
            LIRWSPDInf.Style("left") = 350 & "px"
            LIRWSPDInf.NavigateUrl = "http://10.245.1.6/IRW/INQ_SPDIRWlIST.aspx?pUserID=" & wUserID & "&pSize=" & xSize & "&pFamily=" & xFamily & "&pSlider=" & DSlider1.Text

            LIRWSPDInf1.Style("left") = 450 & "px"
            LIRWSPDInf1.NavigateUrl = "http://10.245.1.6/IRW/INQ_SPDIRWlIST.aspx?pUserID=" & wUserID & "&pSize=" & xSize & "&pFamily=" & xFamily & "&pSlider=" & DSlider2.Text

            LSPDNOURL.Text = dtItemRegisterSheet.Rows(0).Item("SPDNO").ToString.Trim
            LSPDNOURL.NavigateUrl = dtItemRegisterSheet.Rows(0).Item("SPDURL").ToString.Trim

            DSPDNOUrl.Text = dtItemRegisterSheet.Rows(0).Item("SPDURL").ToString.Trim
            '
            '---------------------------------------------------
            LEDXLink1.Visible = False
            LEDXLink2.Visible = False
            '
            LISIPInf.Style("left") = 550 & "px"
            LISIPInf.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pSize=" & xSize & "&pFamily=" & xFamily & "&pSlider=" & DSlider1.Text & "&pBuyer=" & DBuyerCode.Text & "&pSource=IRW"
            '
            LISIPInf1.Style("left") = 650 & "px"
            LISIPInf1.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pSize=" & xSize & "&pFamily=" & xFamily & "&pSlider=" & DSlider2.Text & "&pBuyer=" & DBuyerCode.Text & "&pSource=IRW"
            '
            If InStr(DSPDNOUrl.Text, "NOT FOUND") <= 0 And (InStr(DSPDNOUrl.Text, "(E)SLD-") > 0 Or InStr(DSPDNOUrl.Text, "(E-M)SLD-") > 0) Then
                'SLIDER-1
                If DSlider1.Text <> "" Then
                    str = ""
                    xKey = ""
                    If InStr(DSPDNOUrl.Text, "(E)SLD-1") > 0 Then
                        xKey = "(E)SLD-1"
                    End If
                    If InStr(DSPDNOUrl.Text, "(E-M)SLD-1") > 0 Then
                        str = Mid(DSPDNOUrl.Text, InStr(DSPDNOUrl.Text, "(E-M)SLD-1") + 10, 999)
                        str = Mid(str, InStr(str, "["), InStr(str, "]"))
                        If InStr(str, "/") <= 0 Then
                            xKey = "(E-M)SLD-1"
                        End If
                    End If
                    '
                    If InStr(DSPDNOUrl.Text, xKey) > 0 Then
                        str = Mid(DSPDNOUrl.Text, InStr(DSPDNOUrl.Text, xKey) + Len(xKey), 999)
                        str = Mid(str, InStr(str, "["), InStr(str, "]"))
                        '[QC-NK296-DM]
                        str = Replace(str, "-B", "")
                        str = Replace(str, "SK", "")
                        str = Replace(str, "[", "")
                        str = Replace(str, "]", "")
                        str = Trim(str) & "-"
                        iSlider = str.ToString.Split("-")
                        '
                        If UBound(iSlider) > 2 And iSlider(1) <> "" Then
                            LEDXLink1.NavigateUrl = "http://10.245.1.6/IRW/EDXList01.aspx" & _
                                                    "?&pSize=" & DSize.Text & _
                                                    "&pPuller=" & iSlider(1) & _
                                                    "&pColor=" & iSlider(2)
                            LEDXLink1.Visible = True
                        End If
                    End If
                End If
                'SLIDER-2
                If DSlider2.Text <> "" Then
                    str = ""
                    xKey = ""
                    If InStr(DSPDNOUrl.Text, "(E)SLD-2") > 0 Then
                        xKey = "(E)SLD-2"
                    End If
                    If InStr(DSPDNOUrl.Text, "(E-M)SLD-2") > 0 Then
                        str = Mid(DSPDNOUrl.Text, InStr(DSPDNOUrl.Text, "(E-M)SLD-2") + 10, 999)
                        str = Mid(str, InStr(str, "["), InStr(str, "]"))
                        If InStr(str, "/") <= 0 Then
                            xKey = "(E-M)SLD-2"
                        End If
                    End If
                    '
                    If InStr(DSPDNOUrl.Text, xKey) > 0 Then
                        str = Mid(DSPDNOUrl.Text, InStr(DSPDNOUrl.Text, xKey) + Len(xKey), 999)
                        str = Mid(str, InStr(str, "["), InStr(str, "]"))
                        '[QC-NK296-DM]
                        str = Replace(str, "-B", "")
                        str = Replace(str, "SK", "")
                        str = Replace(str, "[", "")
                        str = Replace(str, "]", "")
                        str = Trim(str) & "-"
                        iSlider = str.ToString.Split("-")
                        '
                        If UBound(iSlider) > 2 And iSlider(1) <> "" Then
                            LEDXLink2.NavigateUrl = "http://10.245.1.6/IRW/EDXList01.aspx" & _
                                                    "?&pSize=" & DSize.Text & _
                                                    "&pPuller=" & iSlider(1) & _
                                                    "&pColor=" & iSlider(2)
                            LEDXLink2.Visible = True
                        End If
                    End If
                End If
            Else
                'SLIDER-1
                If DSlider1.Text <> "" Then
                    LEDXLink1.NavigateUrl = "http://10.245.1.6/IRW/EDXList01.aspx" & _
                                            "?&pSize=" & DSize.Text & _
                                            "&pPuller=" & DSlider1.Text & _
                                            "&pColor=" & ""
                    LEDXLink1.Visible = True
                End If
                'SLIDER-2
                If DSlider2.Text <> "" Then
                    LEDXLink2.NavigateUrl = "http://10.245.1.6/IRW/EDXList01.aspx" & _
                                            "?&pSize=" & DSize.Text & _
                                            "&pPuller=" & DSlider2.Text & _
                                            "&pColor=" & ""
                    LEDXLink2.Visible = True
                End If
            End If
        End If
    End Sub
End Class
