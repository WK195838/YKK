Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class AdminMenu
    Inherits System.Web.UI.Page
    '
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New WAVES.CommonService
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "AdminMenu.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RegisterHistory)
    '**  登錄履歷
    '**
    '*****************************************************************
    Protected Sub BRegisterHistory_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRegisterHistory.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/MagApplyHistory.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Drop)
    '**  Drop
    '**
    '*****************************************************************
    Protected Sub BDrop_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDrop.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/Drop.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ManyBuyer)
    '**  ManyBuyer
    '**
    '*****************************************************************
    Protected Sub BManyBuyer_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BManyBuyer.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ManyBuyer.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SharePuller)
    '**  SharePuller
    '**
    '*****************************************************************
    Protected Sub BSharePuller_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSharePuller.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/SharePuller.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ColorJump)
    '**  ColorJump
    '**
    '*****************************************************************
    Protected Sub BColorJump_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BColorJump.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ColorJump.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(WIP)
    '**  WIP
    '**
    '*****************************************************************
    Protected Sub BWIP_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BWIP.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/WIPList.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RegisterAuthority)
    '**  RegisterAuthority
    '**
    '*****************************************************************
    Protected Sub BRegisterAuthority_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRegisterAuthority.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/RegisterAuthority.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ISOS2ISIP)
    '**  ISOS2ISIP
    '**
    '*****************************************************************
    Protected Sub BISOS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BISOS.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.0.3/ISOS/ISOS2ISIP.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DeplicatPuller)
    '**  DeplicatPuller
    '**
    '*****************************************************************
    Protected Sub BDeplicatPuller_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDeplicatPuller.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/DeplicatPuller.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(NonRubber)
    '**  NonRubber
    '**
    '*****************************************************************
    Protected Sub BNonRubber_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BNonRubber.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NonRubber.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ManySupplier)
    '**  ManySupplier
    '**
    '*****************************************************************
    Protected Sub BManySupplier_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BManySupplier.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ManySupplier.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(NextColorNo)
    '**  NextColorNo
    '**
    '*****************************************************************
    Protected Sub BNextColorNo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BNextColorNo.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NextColorNo.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShortPuller)
    '**  ShortPuller
    '**
    '*****************************************************************
    Protected Sub BShortPuller_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BShortPuller.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ShortPuller.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AutoDRR)
    '**  AutoDRR
    '**
    '*****************************************************************
    Protected Sub BAutoDRR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BAutoDRR.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/AutoDRR.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BodyList)
    '**  BodyList
    '**
    '*****************************************************************
    Protected Sub BBodyList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBodyList.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/BodyList.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BackupList)
    '**  BackupList
    '**
    '*****************************************************************
    Protected Sub BBackupList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBackupList.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/BackupList.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ISIP2RD)
    '**  ISIP2RD
    '**
    '*****************************************************************
    Protected Sub BISIP2RD_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BISIP2RD.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ISIP2RD.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ISIP2EDX)
    '**  ISIP2EDX
    '**
    '*****************************************************************
    Protected Sub BISIP2EDX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BISIP2EDX.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ISIP2EDX.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(NoOrderReceiver)
    '**  NoOrderReceiver
    '**
    '*****************************************************************
    Protected Sub BNoOrderReceiver_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BNoOrderReceiver.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NoOrderReceiver.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ISIP2IRW)
    '**  ISIP2IRW
    '**
    '*****************************************************************
    Protected Sub BISIP2IRW_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BISIP2IRW.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ISIP2IRW.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ISIPNotRegister)
    '**  ISIPNotRegister
    '**
    '*****************************************************************
    Protected Sub BISIPNotRegister_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BISIPNotRegister.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ISIPNotRegister.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(NotUsedColor)
    '**  NotUsedColor
    '**
    '*****************************************************************
    Protected Sub BNotUsedColor_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BNotUsedColor.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NotUsedColor.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(StatusRebuild)
    '**  StatusRebuild
    '**
    '*****************************************************************
    Protected Sub BStatusRebuild_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BStatusRebuild.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/StatusRebuild.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BQCEDXList)
    '**  QCEDXList
    '**
    '*****************************************************************
    Protected Sub BQCEDXList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BQCEDXList.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/QCEDXList.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSPD2EDX)
    '**  SPD2EDX
    '**
    '*****************************************************************
    Protected Sub BSPD2EDX_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSPD2EDX.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/SPD2EDX.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSPDStatusRebuild)
    '**  SPDStatusRebuild
    '**
    '*****************************************************************
    Protected Sub BSPDStatusRebuild_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSPDStatusRebuild.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/SPDStatusRebuild.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSPDMMS2ISIP)
    '**  SPDMMS2ISIP
    '**
    '*****************************************************************
    Protected Sub BSPDMMS2ISIP_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSPDMMS2ISIP.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/SPDMMS2ISIP.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BPullerHistoryList)
    '**  PullerHistoryList
    '**
    '*****************************************************************
    Protected Sub BPullerHistoryList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPullerHistoryList.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/PullerHistoryList.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ColorNoDuplicate)
    '**  ColorNoDuplicate
    '**
    '*****************************************************************
    Protected Sub BColorNoDuplicate_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BColorNoDuplicate.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/ColorNoDuplicate.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(OrderAmount)
    '**  OrderAmount
    '**
    '*****************************************************************
    Protected Sub BOrderAmount_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BOrderAmount.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/OrderAmount.aspx?pUserID=" & UserID & "&pFun=ALL','',''); " + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
End Class
