Public Class SliderPicker
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DQASheet As System.Web.UI.WebControls.Image
    Protected WithEvents BClose As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BDown As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSlider1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSlider2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '不是PostBack
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
            Response.Write(YKK.ShowMessage("所選擇內容已超過上限長度"))
        End If
    End Sub

    Private Sub BClose_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BClose.Click
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();</script>", "Form1.DSliderCode", DContent.Text, "Form1.DSliderGRCode", DSlider2.Text))

    End Sub

End Class
