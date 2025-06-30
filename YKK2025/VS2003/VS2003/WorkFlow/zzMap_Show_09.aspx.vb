Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents OK As System.Web.UI.WebControls.Button
    Protected WithEvents Save As System.Web.UI.WebControls.Button
    Protected WithEvents NG As System.Web.UI.WebControls.Button
    Protected WithEvents DORIMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Bef_OP_Link As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Bef_Map_Link As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBackground As System.Web.UI.WebControls.TextBox
    Protected WithEvents Dsize As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DChainType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBody As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSurface As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCramper As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFrontBack As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFrontBackASS As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterialDetail As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMaterial As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DRefMapFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DDescription As System.Web.UI.WebControls.TextBox

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

    Private Sub DDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDate.TextChanged

    End Sub
End Class
