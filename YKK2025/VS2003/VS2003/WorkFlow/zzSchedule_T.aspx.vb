Public Class Schedule_T
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents DEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents DStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents DCTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents DStep As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents DFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label7 As System.Web.UI.WebControls.Label
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim oSchedule As Object
        Dim pFormNo, pLevel As String
        Dim pStep, pFormSno, pSeqno As Integer
        Dim pCTime As DateTime
        Dim RtnCode As Integer

        pFormNo = DFormNo.Text
        pFormSno = CInt(DFormSno.Text)
        pLevel = DLevel.Text

        pStep = CInt(DStep.Text)
        pSeqno = 1
        pCTime = CDate(DCTime.Text)

        oSchedule = Server.CreateObject("Schedule.WFSchedule")
        RtnCode = oSchedule.AdjustSchedule(pFormNo, pFormSno, pStep, pSeqno, pCTime, pLevel)

        If RtnCode = 1 Then
            DStartTime.Text = "Error"
            DEndTime.Text = "Error"
        Else
            'DStartTime.Text = pStartTime
            'DEndTime.Text = pEndTime
        End If


    End Sub
End Class
