Public Class GetLeadTime_T
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents DStep As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents DFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents DCTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents DStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim oGetLeadTime As Object
        Dim pFormNo, pLevel As String
        Dim pStep As Integer
        Dim pCTime, pStartTime, pEndTime As DateTime
        Dim RtnCode As Integer

        pFormNo = DFormNo.Text
        pStep = CInt(DStep.Text)
        pCTime = CDate(DCTime.Text)
        pLevel = DLevel.Text

        oGetLeadTime = Server.CreateObject("GetLeadTime.WFOPInf")
        oGetLeadTime.pFormNo = pFormNo
        oGetLeadTime.pStep = pStep
        oGetLeadTime.pLevel = pLevel
        RtnCode = oGetLeadTime.LeadTime(pCTime, pStartTime, pEndTime)

        If RtnCode = 1 Then
            DStartTime.Text = "Error"
            DEndTime.Text = "Error"
        Else
            DStartTime.Text = pStartTime
            DEndTime.Text = pEndTime
        End If
    End Sub
End Class
