Public Class SliderPicker
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DQASheet As System.Web.UI.WebControls.Image
    Protected WithEvents BClose As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BDown As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSlider1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSlider2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     �ۭq�[���w
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '���OPostBack
        End If
    End Sub

    Private Sub BDown_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDown.Click
        Dim Str As String = "[" & DSlider1.Text & DSlider2.Text & "]"
        If Len(DContent.Text) + Len(Str) < 255 Then
            If DContent.Text = "" Then
                DContent.Text = Str
            Else
                DContent.Text = DContent.Text + ";" + Str
            End If
        Else
            Response.Write(YKK.ShowMessage("�ҿ�ܤ��e�w�W�L�W������"))
        End If
    End Sub

    Private Sub BClose_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BClose.Click
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();</script>", "Form1.DSliderCode", DContent.Text, "Form1.DSliderGRCode", DSlider2.Text))

    End Sub

End Class
