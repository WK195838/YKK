<%@ Page Language="vb" AutoEventWireup="false" Codebehind="3S_InqCommission.aspx.vb" Inherits="SPD._3S_InqCommission"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>調閱資料</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV style="Z-INDEX: 100; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
				title="篩選項目：" ms_positioning="FlowLayout">篩選項目：</DIV> <!--EndFragment--><asp:imagebutton style="Z-INDEX: 122; POSITION: absolute; TOP: 56px; LEFT: 400px" id="BExcel" runat="server"
				Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:dropdownlist style="Z-INDEX: 121; POSITION: absolute; TOP: 8px; LEFT: 352px" id="DProgress" runat="server"
				Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">開發中</asp:ListItem>
				<asp:ListItem Value="2">開發完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 120; POSITION: absolute; TOP: 8px; LEFT: 480px" id="DSts" runat="server"
				Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">OK</asp:ListItem>
				<asp:ListItem Value="2">NG</asp:ListItem>
				<asp:ListItem Value="3">取消</asp:ListItem>
			</asp:dropdownlist><asp:textbox style="Z-INDEX: 116; POSITION: absolute; TOP: 32px; LEFT: 96px" id="DNo" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:datagrid style="Z-INDEX: 114; POSITION: absolute; TOP: 88px; LEFT: 8px" id="DataGrid1" runat="server"
				Height="300px" Width="950px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="15">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1" HeaderText="案件No">
						<HeaderStyle Width="100px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="案件狀態">
						<HeaderStyle Width="100px" ></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="委託者">
						<HeaderStyle Width="100px" ></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="委託日">
						<HeaderStyle Width="100px" ></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="備註">
						<HeaderStyle Width="350px" ></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field6" ReadOnly="True" HeaderText="審理中工程" HeaderStyle-HorizontalAlign="Center">
						<HeaderStyle Width="200px" ></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="履歷">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:textbox style="Z-INDEX: 111; POSITION: absolute; TOP: 32px; LEFT: 352px" id="DPerson" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:dropdownlist style="Z-INDEX: 102; POSITION: absolute; TOP: 32px; LEFT: 224px" id="DDivision"
				runat="server" Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 103; POSITION: absolute; TOP: 8px; LEFT: 96px" id="DFormName" runat="server"
				Height="40px" Width="248px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist><asp:button style="Z-INDEX: 104; POSITION: absolute; TOP: 56px; LEFT: 320px" id="BEDate" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 105; POSITION: absolute; TOP: 56px; LEFT: 232px" id="DEDate" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox>
			<DIV style="Z-INDEX: 106; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 56px; LEFT: 208px"
				title="" ms_positioning="FlowLayout">～</DIV>
			<asp:button style="Z-INDEX: 107; POSITION: absolute; TOP: 56px; LEFT: 184px" id="BSDate" runat="server"
				Height="20px" Width="24px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 108; POSITION: absolute; TOP: 56px; LEFT: 96px" id="DSDate" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox>
			<DIV style="Z-INDEX: 109; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 56px; LEFT: 8px"
				title="完成日期：" ms_positioning="FlowLayout">完成日期：</DIV>
			<asp:button style="Z-INDEX: 110; POSITION: absolute; TOP: 56px; LEFT: 352px" id="Go" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button></form>
	</body>
</HTML>
