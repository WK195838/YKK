Public Class NextGate_T
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents DStep As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents DUserID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DApplyID As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents DNextStep As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents DCount As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label7 As System.Web.UI.WebControls.Label
    Protected WithEvents DNextGate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNextGate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNextGate3 As System.Web.UI.WebControls.TextBox

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
        Dim oNextGate As Object
        Dim pFormNo As String
        Dim pStep As Integer
        Dim pUserID, pApplyID As String
        Dim RtnCode As Integer
        Dim pNextGate(10) As String
        Dim pNextStep As Integer
        Dim pFlowType As Integer = 0    '0=通知
        Dim pCount As Integer

        pFormNo = DFormNo.Text
        pStep = CInt(DStep.Text)
        pUserID = DUserID.Text
        pApplyID = DApplyID.Text

        oNextGate = Server.CreateObject("NextGate.WFNextGate")
        oNextGate.pFormNo = pFormNo
        oNextGate.pStep = pStep
        oNextGate.pUserID = pUserID
        oNextGate.pApplyID = pApplyID
        RtnCode = oNextGate.NextGate(pNextStep, pNextGate, pCount, pFlowType)  '下一工程的, 號碼, 擔當者, 人數, 處理方法 

        If RtnCode = 1 Then
            DNextStep.Text = "Error"
        Else
            DNextStep.Text = CStr(pNextStep)
            DCount.Text = CStr(pCount)
            DNextGate1.Text = pNextGate(1)
            DNextGate2.Text = pNextGate(2)
            DNextGate3.Text = pNextGate(3)
        End If

    End Sub
End Class
