Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
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

    Private Sub DDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDate.TextChanged

    End Sub
End Class
