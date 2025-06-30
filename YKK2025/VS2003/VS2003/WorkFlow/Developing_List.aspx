<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Developing_List.aspx.vb" Inherits="SPD.Developing_List"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Developing_List</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 96px" runat="server"
					AllowPaging="True" BackColor="White" Width="960px" Height="300px" PageSize="15" BorderStyle="None"
					Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Sts" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Progress" ReadOnly="True" HeaderText="進度">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="委託單">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
							DataTextField="No" HeaderText="No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="SliderCode" ReadOnly="True" HeaderText="Code">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepName" ReadOnly="True" HeaderText="待處理工程">
							<HeaderStyle Width="120px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="工程擔當">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BEndTime" ReadOnly="True" HeaderText="完成預定">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="依賴者">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ApplyTime" ReadOnly="True" HeaderText="依賴日">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
							HeaderText="履歷">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:dropdownlist id="DWDivision" style="Z-INDEX: 115; LEFT: 616px; POSITION: absolute; TOP: 16px"
					runat="server" BackColor="Yellow" Width="120px" Height="40px" ForeColor="Blue">
					<asp:ListItem Value="3" Selected="True">０～３日</asp:ListItem>
					<asp:ListItem Value="6">４～６日</asp:ListItem>
					<asp:ListItem Value="10">７～１０日</asp:ListItem>
					<asp:ListItem Value="999">１０～９９９日</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDays" style="Z-INDEX: 114; LEFT: 744px; POSITION: absolute; TOP: 16px" runat="server"
					BackColor="Yellow" Width="120px" Height="40px" ForeColor="Blue">
					<asp:ListItem Value="3" Selected="True">０～３日</asp:ListItem>
					<asp:ListItem Value="6">４～６日</asp:ListItem>
					<asp:ListItem Value="10">７～１０日</asp:ListItem>
					<asp:ListItem Value="999">１０～９９９日</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DProgress" style="Z-INDEX: 113; LEFT: 488px; POSITION: absolute; TOP: 16px"
					runat="server" BackColor="Yellow" Width="120px" Height="40px" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="正常">正常</asp:ListItem>
					<asp:ListItem Value="延遲" Selected="True">延遲</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DBuyer" style="Z-INDEX: 112; LEFT: 320px; POSITION: absolute; TOP: 64px" runat="server"
					BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ForeColor="Blue">DBuyer</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 109; LEFT: 448px; POSITION: absolute; TOP: 64px"
					runat="server" BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ForeColor="Blue">DSliderCode</asp:textbox><asp:imagebutton id="BExcel" style="Z-INDEX: 111; LEFT: 632px; POSITION: absolute; TOP: 64px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:textbox id="DEndStep" style="Z-INDEX: 108; LEFT: 344px; POSITION: absolute; TOP: 40px" runat="server"
					BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ForeColor="Blue">DEndStep</asp:textbox><asp:textbox id="DStartStep" style="Z-INDEX: 107; LEFT: 192px; POSITION: absolute; TOP: 40px"
					runat="server" BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ForeColor="Blue">DStartStep</asp:textbox><asp:button id="Go" style="Z-INDEX: 106; LEFT: 584px; POSITION: absolute; TOP: 64px" runat="server"
					BackColor="White" Width="40px" Height="24px" ForeColor="Blue" Text="Go"></asp:button><asp:textbox id="DApplyName" style="Z-INDEX: 105; LEFT: 192px; POSITION: absolute; TOP: 64px"
					runat="server" BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ForeColor="Blue">DApplyName</asp:textbox><asp:dropdownlist id="DSts" style="Z-INDEX: 104; LEFT: 360px; POSITION: absolute; TOP: 16px" runat="server"
					BackColor="Yellow" Width="120px" Height="40px" ForeColor="Blue">
					<asp:ListItem Value="1" Selected="True">開發中</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 103; LEFT: 192px; POSITION: absolute; TOP: 16px"
					runat="server" Height="40px" Width="165px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:textbox id="DRefresh" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					BorderStyle="None" Height="24px" Width="174px" BackColor="White" ForeColor="Blue" ReadOnly="True">DDateTime</asp:textbox><asp:textbox id="DDateTime" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 16px" runat="server"
					BorderStyle="None" Height="24px" Width="174px" BackColor="White" ForeColor="Blue" ReadOnly="True">DDateTime</asp:textbox><asp:label id="Label1" style="Z-INDEX: 110; LEFT: 320px; POSITION: absolute; TOP: 40px" runat="server"
					Height="8px" Width="24px" ForeColor="Blue">～</asp:label></FONT></form>
	</body>
</HTML>
