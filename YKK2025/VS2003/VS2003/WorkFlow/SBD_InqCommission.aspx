<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SBD_InqCommission.aspx.vb" Inherits="SPD.SBD_InqCommission"%>
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
			<DIV style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				title="篩選項目：" ms_positioning="FlowLayout">篩選項目：</DIV> <!--EndFragment-->
			<asp:dropdownlist id="DType" style="Z-INDEX: 121; LEFT: 352px; POSITION: absolute; TOP: 80px" runat="server"
				Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue" EnableViewState="False">
				<asp:ListItem Value="未封存">未封存</asp:ListItem>
				<asp:ListItem Value="已封存">已封存</asp:ListItem>
			</asp:dropdownlist><asp:imagebutton style="Z-INDEX: 120; LEFT: 536px; POSITION: absolute; TOP: 80px" id="BExcel" runat="server"
				Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:dropdownlist style="Z-INDEX: 119; LEFT: 352px; POSITION: absolute; TOP: 8px" id="DProgress" runat="server"
				Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">開發中</asp:ListItem>
				<asp:ListItem Value="2">開發完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 118; LEFT: 480px; POSITION: absolute; TOP: 8px" id="DSts" runat="server"
				Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">OK</asp:ListItem>
				<asp:ListItem Value="2">NG</asp:ListItem>
				<asp:ListItem Value="3">取消</asp:ListItem>
			</asp:dropdownlist><asp:textbox style="Z-INDEX: 116; LEFT: 224px; POSITION: absolute; TOP: 56px" id="DDevNo" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" Visible="False" BorderStyle="Groove" ReadOnly="True">DDevNo</asp:textbox><asp:textbox style="Z-INDEX: 115; LEFT: 96px; POSITION: absolute; TOP: 32px" id="DNo" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:datagrid style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 112px" id="DataGrid1" runat="server"
				Height="300px" Width="790px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="15">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="開發狀態">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="工程">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="工程擔當">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field6" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field7" ReadOnly="True" HeaderText="完成預定">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field8" ReadOnly="True" HeaderText="閱讀期限">
						<HeaderStyle Width="90px"></HeaderStyle>
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
			</asp:datagrid><asp:textbox style="Z-INDEX: 113; LEFT: 352px; POSITION: absolute; TOP: 56px" id="DCodeNo" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" Visible="False" BorderStyle="Groove" ReadOnly="True">DCodeNo</asp:textbox><asp:textbox style="Z-INDEX: 112; LEFT: 96px; POSITION: absolute; TOP: 56px" id="DRno" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" Visible="False" BorderStyle="Groove" ReadOnly="True">DRno</asp:textbox><asp:textbox style="Z-INDEX: 111; LEFT: 352px; POSITION: absolute; TOP: 32px" id="DPerson" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:textbox style="Z-INDEX: 101; LEFT: 480px; POSITION: absolute; TOP: 32px" id="DBuyer" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:dropdownlist style="Z-INDEX: 102; LEFT: 224px; POSITION: absolute; TOP: 32px" id="DDivision"
				runat="server" Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 8px" id="DFormName" runat="server"
				Height="40px" Width="248px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist><asp:button style="Z-INDEX: 104; LEFT: 320px; POSITION: absolute; TOP: 80px" id="BEDate" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 105; LEFT: 232px; POSITION: absolute; TOP: 80px" id="DEDate" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DEDate</asp:textbox>
			<DIV style="DISPLAY: inline; Z-INDEX: 106; LEFT: 208px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				title="" ms_positioning="FlowLayout">～</DIV>
			<asp:button style="Z-INDEX: 107; LEFT: 184px; POSITION: absolute; TOP: 80px" id="BSDate" runat="server"
				Height="20px" Width="24px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 108; LEFT: 96px; POSITION: absolute; TOP: 80px" id="DSDate" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DSDate</asp:textbox>
			<DIV style="DISPLAY: inline; Z-INDEX: 109; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				title="完成日期：" ms_positioning="FlowLayout">完成日期：</DIV>
			<asp:button style="Z-INDEX: 110; LEFT: 488px; POSITION: absolute; TOP: 80px" id="Go" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button></form>
	</body>
</HTML>
