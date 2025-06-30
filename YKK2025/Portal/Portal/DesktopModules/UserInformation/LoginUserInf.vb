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
'***********************************************************************************************
'**
'** Namespace : LoginUserInf
'**
'***********************************************************************************************
Namespace LoginUserInf

'***********************************************************************************************
'**
'** Class : PutCookies
'**
'***********************************************************************************************
    Public Class PutCookies
        Inherits PortalModuleBase
        Implements IActionable
        ' This call is required by the Web Form Designer.
        Private Sub InitializeComponent()
        End Sub

        Private Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Init
            InitializeComponent()
        End Sub
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Try
                If Not Page.IsPostBack Then
                    Dim mUser As DotNetNuke.Entities.Users.UserInfo = DotNetNuke.Entities.Users.UserController.GetCurrentUserInfo 
                    Response.Cookies("YKK_PORTAL")("UserID") = mUser.Username


                End If
            Catch ex As Exception
                Exceptions.ProcessModuleLoadException(Me, ex)
            End Try
        End Sub

        Public ReadOnly Property ModuleActions() As ModuleActionCollection Implements IActionable.ModuleActions
            Get
                Dim actions As New ModuleActionCollection
                actions.Add(GetNextActionID(), Localization.GetString(ModuleActionType.AddContent, LocalResourceFile), ModuleActionType.AddContent, "", "", EditUrl(), False, DotNetNuke.Security.SecurityAccessLevel.Edit, False, False)
                Return actions
            End Get
        End Property

    End Class


End Namespace