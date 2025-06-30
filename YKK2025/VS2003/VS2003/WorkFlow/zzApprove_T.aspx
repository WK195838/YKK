<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzApprove_T.aspx.vb" Inherits="SPD.Approve_T"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Approve_T</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:Label id="Label1" style="Z-INDEX: 108; LEFT: 24px; POSITION: absolute; TOP: 24px" runat="server"
					Width="120px" Height="24px">表單號碼：</asp:Label>
				<asp:label id="Label5" style="Z-INDEX: 113; LEFT: 32px; POSITION: absolute; TOP: 152px" runat="server"
					Width="120px" Height="24px">登錄者：</asp:label>
				<asp:label id="Label6" style="Z-INDEX: 110; LEFT: 32px; POSITION: absolute; TOP: 176px" runat="server"
					Width="120px" Height="24px">申請者：</asp:label>
				<asp:textbox id="DUserID" style="Z-INDEX: 111; LEFT: 152px; POSITION: absolute; TOP: 144px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">1</asp:textbox>
				<asp:textbox id="DApplyID" style="Z-INDEX: 112; LEFT: 152px; POSITION: absolute; TOP: 176px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow">1</asp:textbox>
				<asp:Label id="Label4" style="Z-INDEX: 107; LEFT: 32px; POSITION: absolute; TOP: 112px" runat="server"
					Width="120px" Height="24px">序號：</asp:Label>
				<asp:TextBox id="DSeqNo" style="Z-INDEX: 106; LEFT: 152px; POSITION: absolute; TOP: 112px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">1</asp:TextBox>
				<asp:Label id="Label3" style="Z-INDEX: 101; LEFT: 24px; POSITION: absolute; TOP: 84px" runat="server"
					Width="120px" Height="24px">工程代號：</asp:Label>
				<asp:Label id="Label2" style="Z-INDEX: 100; LEFT: 24px; POSITION: absolute; TOP: 54px" runat="server"
					Width="200px" Height="24px">表單流水號：　　</asp:Label>
				<asp:TextBox id="DFormNo" style="Z-INDEX: 102; LEFT: 152px; POSITION: absolute; TOP: 24px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">000001</asp:TextBox>
				<asp:TextBox id="DFormSno" style="Z-INDEX: 103; LEFT: 152px; POSITION: absolute; TOP: 54px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">1</asp:TextBox>
				<asp:TextBox id="DStep" style="Z-INDEX: 104; LEFT: 152px; POSITION: absolute; TOP: 84px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">2</asp:TextBox>
				<asp:Button id="Button1" style="Z-INDEX: 105; LEFT: 128px; POSITION: absolute; TOP: 240px" runat="server"
					Width="128px" Height="40px" Text="Go"></asp:Button></FONT>
		</form>
	</body>
</HTML>
