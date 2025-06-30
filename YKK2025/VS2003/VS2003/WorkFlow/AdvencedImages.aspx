<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AdvencedImages.aspx.vb" Inherits="SPD.AdvencedImages"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AdvencedImages</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<div>
				<!-- ****************************************************************************************** -->
				<!-- ** Button                                                                                  -->
				<!-- ****************************************************************************************** -->
				<!-- Find -->
				<asp:Button ID="BFind" runat="server" Height="64px" Style="Z-INDEX: 104; POSITION: absolute; TOP: 9px; LEFT: 376px"
					Text="Go" Width="96px" Font-Bold="True" Font-Size="Larger" />
				<!-- ****************************************************************************************** -->
				<!-- ** Puller Key                                                                              -->
				<!-- ****************************************************************************************** -->
				<asp:TextBox ID="TextBox2" runat="server" BackColor="WHITE" BorderStyle="None" ForeColor="BLACK"
					Height="18px" ReadOnly="True" Style="Z-INDEX: 126; POSITION: absolute; TEXT-ALIGN: left; TOP: 16px; LEFT: 8px"
					Width="352px" Font-Bold="True" Font-Size="Larger">Sample: AD700 ~ AD999</asp:TextBox>
				<asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
					Height="18px" ReadOnly="True" Style="Z-INDEX: 126; POSITION: absolute; TEXT-ALIGN: left; TOP: 48px; LEFT: 8px"
					Width="120px" Font-Bold="True" Font-Size="Larger">Start ~ End</asp:TextBox>
				<asp:TextBox ID="DKStart" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
					Height="18px" Style="Z-INDEX: 126; POSITION: absolute; TEXT-ALIGN: left; TOP: 48px; LEFT: 128px"
					Width="100px" Font-Bold="True" Font-Size="Larger"></asp:TextBox>
				<asp:TextBox ID="TextBox11" runat="server" BackColor="white" BorderStyle="None" ForeColor="black"
					Height="18px" ReadOnly="True" Style="Z-INDEX: 126; POSITION: absolute; TEXT-ALIGN: left; TOP: 48px; LEFT: 240px"
					Width="32px" Font-Bold="True" Font-Size="Larger"> ~ </asp:TextBox>
				<asp:TextBox ID="DKEnd" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
					Height="18px" Style="Z-INDEX: 126; POSITION: absolute; TEXT-ALIGN: left; TOP: 48px; LEFT: 264px"
					Width="100px" Font-Bold="True" Font-Size="Larger"></asp:TextBox>
				<!-- ****************************************************************************************** -->
				<!-- ** DataList                                                                          -->
				<!-- ****************************************************************************************** -->
				<asp:datalist id="DataList1" style="Z-INDEX: 105; POSITION: absolute; TOP: 128px; LEFT: 0px" runat="server"
					Width="10px" Height="93px" CellSpacing="2" RepeatColumns="5" RepeatDirection="Horizontal">
					<ItemTemplate>
						<asp:Table Width="100%" GridLines="Both" Font-Size="8pt" Runat="server" ID="Table2">
							<asp:TableRow>
								<asp:TableCell HorizontalAlign="Left" BorderWidth="1" BorderStyle="Solid" BorderColor="#990000"
									Font-Size="10" BackColor="#ccff66">
									<%#Container.DataItem("Code")%>
								</asp:TableCell>
							</asp:TableRow>
							<asp:TableRow>
								<asp:TableCell HorizontalAlign="Left" BorderWidth="1" BorderStyle="Solid" BorderColor="#990000"
									Font-Size="10" BackColor="#ccff66">
									<%# Container.DataItem("No") %>
								</asp:TableCell>
							</asp:TableRow>
							<asp:TableRow>
								<asp:TableCell BorderWidth="1" BorderStyle="Solid" BorderColor="#990000" HorizontalAlign="Center"
									BackColor="#ffff66">
									<asp:Image ID="Image1" Runat="server" Height="230" Width="200" ImageUrl='<%# Container.DataItem("ImagePath") %>' >
									</asp:Image>
								</asp:TableCell>
							</asp:TableRow>
						</asp:Table>
					</ItemTemplate>
				</asp:datalist>
				<!-- ****************************************************************************************** -->
			</div>
		</form>
	</body>
</HTML>
