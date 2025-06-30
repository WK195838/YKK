' Code adapted from DAL Builder Pro from www.dotnetnuke.dk

Imports System
Imports System.Drawing
Imports System.Web.UI.WebControls
Imports DotNetNuke
Imports DotNetNuke.Common
Imports DotNetNuke.Services.Search
Imports DotNetNuke.Entities.Modules
Imports DotNetNuke.Entities.Modules.Actions
Imports DotNetNuke.Services.Localization
Imports DotNetNuke.Services.Exceptions

Namespace ImportantMessage

    Public Class ShowMessage
        Inherits PortalModuleBase
        Implements IActionable
        ' This call is required by the Web Form Designer.
        Private Sub InitializeComponent()
        End Sub

        Private Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Init
            InitializeComponent()
        End Sub
        Protected WithEvents lblHelloWorld As System.Web.UI.WebControls.Label

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Try
                If Not Page.IsPostBack Then
                    Dim helloText As String = CType(Settings("helloText"), String)
                    If (helloText Is Nothing) Then
                        helloText = "Hello World"
                    End If
                    lblHelloWorld.Text = helloText
                End If
            Catch ex As Exception
                Exceptions.ProcessModuleLoadException(Me, ex)
            End Try
        End Sub
        Public ReadOnly Property ModuleActions() As ModuleActionCollection Implements IActionable.ModuleActions
            Get
                Dim actions As New ModuleActionCollection
                actions.Add(GetNextActionID(), Localization.GetString(ModuleActionType.AddContent, LocalResourceFile), ModuleActionType.AddContent, "", "", EditUrl(), False, DotNetNuke.Security.SecurityAccessLevel.Edit, True, False)
                Return actions
            End Get
        End Property
    End Class

    Public Class WriteMessage
	Inherits PortalModuleBase

        Protected WithEvents plEditText As DotNetNuke.UI.UserControls.LabelControl
        Protected WithEvents txtEditText As System.Web.UI.WebControls.TextBox
        Protected WithEvents cmdUpdate As System.Web.UI.WebControls.LinkButton
        Protected WithEvents cmdCancel As System.Web.UI.WebControls.LinkButton
        ' This call is required by the Web Form Designer.
        Private Sub InitializeComponent()
        End Sub

        Private Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Init
            InitializeComponent()
        End Sub
        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Not Page.IsPostBack Then
                Dim helloText As String = CType(Settings("helloText"), String)
                If (helloText Is Nothing) Then
                    helloText = "Hello World"
                End If
                txtEditText.Text = helloText
            End If
        End Sub

        Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
            Response.Redirect(Globals.NavigateURL(), True)
        End Sub

        Private Sub cmdUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
            Dim moduleCtrl As New ModuleController
            moduleCtrl.UpdateModuleSetting(ModuleId, "helloText", txtEditText.Text)
            Response.Redirect(Globals.NavigateURL(), True)
        End Sub

    End Class


End Namespace