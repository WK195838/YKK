<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UPDocument.aspx.vb" Inherits="SPD.UPDocument"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>上傳文件</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DUPDocument" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\UPDocument_001.jpg" Width="531px" Height="258px"></asp:image>
				<asp:RequiredFieldValidator id="DDocFileNameRqd" style="Z-INDEX: 113; LEFT: 16px; POSITION: absolute; TOP: 256px"
					runat="server" Width="408px" Height="24px" ErrorMessage="異常：需指定上傳檔案" Display="Dynamic" ControlToValidate="DDocFileName"></asp:RequiredFieldValidator><asp:dropdownlist id="DMaker" style="Z-INDEX: 110; LEFT: 104px; POSITION: absolute; TOP: 189px" runat="server"
					Width="128px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="0" Selected="True">月</asp:ListItem>
					<asp:ListItem Value="1">年</asp:ListItem>
				</asp:dropdownlist><asp:button id="BSave" style="Z-INDEX: 109; LEFT: 432px; POSITION: absolute; TOP: 256px" runat="server"
					Width="105px" Height="32px" Text="儲存"></asp:button><asp:dropdownlist id="DClass" style="Z-INDEX: 108; LEFT: 104px; POSITION: absolute; TOP: 88px" runat="server"
					Width="168px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="0" Selected="True">月</asp:ListItem>
					<asp:ListItem Value="1">年</asp:ListItem>
				</asp:dropdownlist><asp:button id="BMakerDate" style="Z-INDEX: 107; LEFT: 508px; POSITION: absolute; TOP: 189px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DMakerTime" style="Z-INDEX: 106; LEFT: 336px; POSITION: absolute; TOP: 189px"
					runat="server" Width="172px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DMakerTime</asp:textbox><asp:dropdownlist id="DMonth" style="Z-INDEX: 105; LEFT: 440px; POSITION: absolute; TOP: 122px" runat="server"
					Width="92px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="01" Selected="True">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
					<asp:ListItem Value="03">03</asp:ListItem>
					<asp:ListItem Value="04">04</asp:ListItem>
					<asp:ListItem Value="05">05</asp:ListItem>
					<asp:ListItem Value="06">06</asp:ListItem>
					<asp:ListItem Value="07">07</asp:ListItem>
					<asp:ListItem Value="08">08</asp:ListItem>
					<asp:ListItem Value="09">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DYear" style="Z-INDEX: 104; LEFT: 288px; POSITION: absolute; TOP: 122px" runat="server"
					Width="92px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="2005" Selected="True">2005</asp:ListItem>
					<asp:ListItem Value="2006">2006</asp:ListItem>
					<asp:ListItem Value="2007">2007</asp:ListItem>
					<asp:ListItem Value="2008">2008</asp:ListItem>
					<asp:ListItem Value="2009">2009</asp:ListItem>
					<asp:ListItem Value="2010">2010</asp:ListItem>
					<asp:ListItem Value="2011">2011</asp:ListItem>
					<asp:ListItem Value="2012">2012</asp:ListItem>
					<asp:ListItem Value="2013">2013</asp:ListItem>
					<asp:ListItem Value="2014">2014</asp:ListItem>
					<asp:ListItem Value="2015">2015</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DDescription" style="Z-INDEX: 103; LEFT: 104px; POSITION: absolute; TOP: 156px"
					runat="server" Width="426px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DDescription</asp:textbox><asp:dropdownlist id="DType" style="Z-INDEX: 102; LEFT: 104px; POSITION: absolute; TOP: 122px" runat="server"
					Width="128px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="0" Selected="True">月報</asp:ListItem>
					<asp:ListItem Value="1">年報</asp:ListItem>
				</asp:dropdownlist><INPUT id="DDocFileName" style="Z-INDEX: 101; LEFT: 104px; WIDTH: 424px; POSITION: absolute; TOP: 224px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="51" name="File1" runat="server"></FONT>
			<asp:RequiredFieldValidator id="DDescriptionRqd" style="Z-INDEX: 112; LEFT: 16px; POSITION: absolute; TOP: 256px"
				runat="server" Width="408px" Height="24px" ErrorMessage="異常：需輸入簡介" Display="Dynamic" ControlToValidate="DDescription"></asp:RequiredFieldValidator></form>
	</body>
</HTML>
