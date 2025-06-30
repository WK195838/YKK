Public Class SufaceSuppilerSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOrderTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCLT As System.Web.UI.WebControls.TextBox
    Protected WithEvents LQCCheck6File As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BQCDate1 As System.Web.UI.WebControls.Button
    Protected WithEvents BQCDate2 As System.Web.UI.WebControls.Button
    Protected WithEvents BQCDate3 As System.Web.UI.WebControls.Button
    Protected WithEvents DQCResult3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCResult2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCRemark3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCDate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEADesc1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEACheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck5 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LFinalSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents DEnglishName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBFinalDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBFinalDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAllowSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DColor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOrderTime As System.Web.UI.WebControls.Button
    Protected WithEvents BReqDelDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSliderSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LQCReqFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LOPManualFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LEACheckFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFinalFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LForcastFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LManufFlowFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAttachSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DORNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DSufaceSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LCustSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DQCCheck6File As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DFinalSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCReqFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DCustSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DOPManualFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFinalFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DContactFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DForcastFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DManufFlowFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DEACheckFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton

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

End Class
