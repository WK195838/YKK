<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzMap_Show_09.aspx.vb" Inherits="SPD.Map_Show_09" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>製圖委託書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">
		BODY { PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 4px; PADDING-TOP: 0px }
		BODY { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		TABLE { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		TR { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		TD { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		UL { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		LI { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		.normal { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		H1 { FONT-WEIGHT: 900; FONT-SIZE: 10.5pt; COLOR: #666666; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		.small { FONT-SIZE: 7.5pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		.error { COLOR: #ff0033 }
		.required { FONT-WEIGHT: 900; COLOR: #ff0033; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		.success { FONT-WEIGHT: 900; FONT-SIZE: 11pt; MARGIN: 10px 0px; COLOR: #009933 }
		</style>
		<script language="javascript" type="text/javascript">
			/// <summary>
			/// Launches the DatePicker page in a popup window, 
			/// passing a JavaScript reference to the field that we want to set.
			/// </summary>
			/// <param name="strField">String. The JavaScript reference to the field that we want to set, in the format: FormName.FieldName
			/// Please note that JavaScript is case-sensitive.</param>
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"><IMG style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px" height="113" alt=""
					src="file:///C:\Inetpub\wwwroot\SSD\Images\MapSheet_001.jpg" width="593"> <INPUT id="DRefMapFile" style="Z-INDEX: 128; LEFT: 162px; WIDTH: 426px; POSITION: absolute; TOP: 435px; HEIGHT: 24px; BACKGROUND-COLOR: #ffff00"
					type="file" size="51" name="File1" runat="server">
				<asp:DropDownList id="DFrontBackASS" style="Z-INDEX: 125; LEFT: 268px; POSITION: absolute; TOP: 368px"
					runat="server" BackColor="Yellow" ForeColor="Blue" Height="24px" Width="323px">
					<asp:ListItem Value="3">一樣</asp:ListItem>
				</asp:DropDownList>
				<asp:TextBox id="DSurface" style="Z-INDEX: 123; LEFT: 393px; POSITION: absolute; TOP: 298px"
					runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px" Width="198px">DSurface</asp:TextBox>
				<asp:Button id="NG" style="Z-INDEX: 111; LEFT: 272px; POSITION: absolute; TOP: 864px" runat="server"
					Height="32px" Width="105px" Text="NG"></asp:Button>
				<asp:Button id="Save" style="Z-INDEX: 110; LEFT: 384px; POSITION: absolute; TOP: 864px" runat="server"
					Height="32px" Width="105px" Text="儲存"></asp:Button><IMG style="Z-INDEX: 108; LEFT: 8px; POSITION: absolute; TOP: 648px" height="110" alt=""
					src="file:///C:\Inetpub\wwwroot\SSD\Images\Sheet_Delivery.jpg" width="593"><IMG style="Z-INDEX: 107; LEFT: 8px; POSITION: absolute; TOP: 776px" height="78" alt=""
					src="file:///C:\Inetpub\wwwroot\SSD\Images\Sheet_Delay.jpg" width="593"><IMG style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 568px" height="77" alt=""
					src="file:///C:\Inetpub\wwwroot\SSD\Images\MapSheet_003.jpg" width="593">
				<asp:Button id="OK" style="Z-INDEX: 109; LEFT: 496px; POSITION: absolute; TOP: 864px" runat="server"
					Height="32px" Width="105px" Text="完成"></asp:Button>
				<asp:HyperLink id="Bef_OP_Link" style="Z-INDEX: 112; LEFT: 168px; POSITION: absolute; TOP: 730px"
					runat="server" Height="16px" Width="104px" Target="_blank" NavigateUrl="aaa">Bef_OP_Link</asp:HyperLink>
				<asp:HyperLink id="Bef_Map_Link" style="Z-INDEX: 113; LEFT: 288px; POSITION: absolute; TOP: 64px"
					runat="server" Height="16px" Width="40px" Target="_blank" NavigateUrl="BoardEdit.aspx">Bef_Map_Link</asp:HyperLink>
				<asp:TextBox id="DNo" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 86px" runat="server"
					Width="96px" Height="25px" ForeColor="Blue" BackColor="Yellow">DNo</asp:TextBox>
				<asp:TextBox id="DORIMapNo" style="Z-INDEX: 103; LEFT: 288px; POSITION: absolute; TOP: 86px"
					runat="server" Width="115px" Height="25px" ForeColor="Blue" BackColor="Yellow">DOriMapNo</asp:TextBox></FONT>
			<asp:TextBox id="DDate" style="Z-INDEX: 104; LEFT: 477px; POSITION: absolute; TOP: 86px" runat="server"
				Width="115px" Height="25px" ForeColor="Blue" BackColor="Yellow">DDate</asp:TextBox>
			<a href="javascript:;" onclick="calendarPicker('Form1.DDate');" title="Pick Date from Calendar">
				pick</a> <IMG style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 120px" height="451" alt=""
				src="file:///C:\Inetpub\wwwroot\SSD\Images\MapSheet_002.jpg" width="593">
			<asp:TextBox id="DBuyer" style="Z-INDEX: 114; LEFT: 162px; POSITION: absolute; TOP: 128px" runat="server"
				BackColor="Yellow" ForeColor="Blue" Height="25px" Width="136px">DBuyer</asp:TextBox>
			<asp:TextBox id="DSellVendor" style="Z-INDEX: 115; LEFT: 456px; POSITION: absolute; TOP: 128px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px" Width="136px">DSellVendor</asp:TextBox>
			<asp:TextBox id="DBackground" style="Z-INDEX: 118; LEFT: 162px; POSITION: absolute; TOP: 196px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px" Width="430px">DBackground</asp:TextBox>
			<asp:TextBox id="DPerson" style="Z-INDEX: 117; LEFT: 456px; POSITION: absolute; TOP: 160px" runat="server"
				BackColor="Yellow" ForeColor="Blue" Height="25px" Width="136px">DPerson</asp:TextBox>
			<asp:TextBox id="DDivision" style="Z-INDEX: 116; LEFT: 162px; POSITION: absolute; TOP: 162px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px" Width="136px">DDivision</asp:TextBox>
			<asp:DropDownList id="Dsize" style="Z-INDEX: 119; LEFT: 58px; POSITION: absolute; TOP: 266px" runat="server"
				BackColor="Yellow" Height="24px" Width="95px" ForeColor="Blue">
				<asp:ListItem Value="3">3</asp:ListItem>
			</asp:DropDownList>
			<asp:DropDownList id="DChainType" style="Z-INDEX: 120; LEFT: 162px; POSITION: absolute; TOP: 266px"
				runat="server" BackColor="Yellow" Height="24px" Width="136px" ForeColor="Blue">
				<asp:ListItem Value="3">CF</asp:ListItem>
			</asp:DropDownList>
			<asp:DropDownList id="DBody" style="Z-INDEX: 121; LEFT: 310px; POSITION: absolute; TOP: 266px" runat="server"
				BackColor="Yellow" Height="24px" Width="282px" ForeColor="Blue">
				<asp:ListItem Value="3">CF</asp:ListItem>
			</asp:DropDownList>
			<asp:TextBox id="DCramper" style="Z-INDEX: 122; LEFT: 162px; POSITION: absolute; TOP: 298px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px" Width="136px">DCramper</asp:TextBox>
			<asp:DropDownList id="DFrontBack" style="Z-INDEX: 124; LEFT: 268px; POSITION: absolute; TOP: 334px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="24px" Width="323px">
				<asp:ListItem Value="3">相同</asp:ListItem>
			</asp:DropDownList>
			<asp:DropDownList id="DMaterial" style="Z-INDEX: 126; LEFT: 162px; POSITION: absolute; TOP: 402px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="24px" Width="95px">
				<asp:ListItem Value="3">3</asp:ListItem>
			</asp:DropDownList>
			<asp:DropDownList id="DMaterialDetail" style="Z-INDEX: 127; LEFT: 264px; POSITION: absolute; TOP: 402px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="24px" Width="327px">
				<asp:ListItem Value="3">CF</asp:ListItem>
			</asp:DropDownList>
			<asp:TextBox id="DDescription" style="Z-INDEX: 129; LEFT: 162px; POSITION: absolute; TOP: 468px"
				runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px" Width="430px">DDescription</asp:TextBox>
		</form>
	</body>
</HTML>
