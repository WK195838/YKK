<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_WorkTime.aspx.vb" Inherits="SPD.HR_WorkTime"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�X�Ԥ��</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 100; LEFT: 216px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Width="168px" Height="40px"></asp:dropdownlist>
				<asp:dropdownlist id="DMonth" style="Z-INDEX: 107; LEFT: 112px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="100px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist>
				<asp:dropdownlist id="DYear" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="100px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:datagrid id="DataGrid1" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					BackColor="White" Width="800px" Height="100px" BorderStyle="None" PageSize="15" Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="DateDesc" ReadOnly="True" HeaderText="��d��">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="NameDesc" ReadOnly="True" HeaderText="�m�W">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="HRWDivName" ReadOnly="True" HeaderText="����">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="JobName" ReadOnly="True" HeaderText="¾��">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TimeA" ReadOnly="True" HeaderText="�W�Z��d">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TimeB" ReadOnly="True" HeaderText="�U�Z��d">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OTURL" DataNavigateUrlFormatString="{0}" DataTextField="OTDesc"
							HeaderText="�[�Z">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="VAURL" DataNavigateUrlFormatString="{0}" DataTextField="VADesc"
							HeaderText="�а�">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="AwayURL" DataNavigateUrlFormatString="{0}"
							DataTextField="AwayDesc" HeaderText="�~�X">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="AworkURL" DataNavigateUrlFormatString="{0}"
							DataTextField="AworkDesc" HeaderText="�ɤu">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:imagebutton id="BExcel" style="Z-INDEX: 103; LEFT: 568px; POSITION: absolute; TOP: 8px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:button id="Go" style="Z-INDEX: 102; LEFT: 520px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="White" ForeColor="Blue" Width="40px" Height="24px" Text="Go"></asp:button><asp:textbox id="DName" style="Z-INDEX: 101; LEFT: 384px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Width="120px" Height="20px" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox></FONT></form>
	</body>
</HTML>
