<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InqCommission.aspx.vb" Inherits="InqCommission" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
	<HEAD>
		<title>調閱資料</title>
	    <script language="javascript" src="RegisterItem.js"></script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 117; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 512px; POSITION: absolute; TOP: 77px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px"></asp:imagebutton><asp:dropdownlist id="DProgress" style="Z-INDEX: 101; LEFT: 342px; POSITION: absolute; TOP: 8px" runat="server"
				Height="26px" Width="120px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">核定中</asp:ListItem>
				<asp:ListItem Value="2">完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DSts" style="Z-INDEX: 102; LEFT: 463px; POSITION: absolute; TOP: 8px" runat="server"
				Height="26px" Width="120px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">OK</asp:ListItem>
				<asp:ListItem Value="3">取消</asp:ListItem>
			</asp:dropdownlist><asp:textbox id="DNo" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 34px" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True" MaxLength="20">DNo</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 104; LEFT: 10px; POSITION: absolute; TOP: 114px" runat="server"
				Height="100px" Width="814px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="20">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="工程">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
                    <asp:BoundColumn DataField="Code" HeaderText="Item Code"></asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="工程擔當">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
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
			</asp:datagrid><asp:textbox id="DName" style="Z-INDEX: 105; LEFT: 222px; POSITION: absolute; TOP: 34px" runat="server"
				Height="20px" Width="114px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True" MaxLength="20">DName</asp:textbox><asp:dropdownlist id="DFormName" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				Height="26px" Width="245px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True" Font-Size="10pt"></asp:dropdownlist><asp:button id="BEDate" style="Z-INDEX: 107; LEFT: 345px; POSITION: absolute; TOP: 77px" runat="server"
				Height="25px" Width="28px" Text="....."></asp:button><asp:textbox id="DEDate" style="Z-INDEX: 108; LEFT: 249px; POSITION: absolute; TOP: 77px" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 118; LEFT: 228px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 82px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:button id="BSDate" style="Z-INDEX: 109; LEFT: 198px; POSITION: absolute; TOP: 77px" runat="server"
				Height="26px" Width="28px" Text="....."></asp:button><asp:textbox id="DSDate" style="Z-INDEX: 110; LEFT: 96px; POSITION: absolute; TOP: 77px" runat="server"
				Height="20px" Width="93px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox>
			<DIV title="申請日期：" style="DISPLAY: inline; Z-INDEX: 119; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 78px; HEIGHT: 24px"
				ms_positioning="FlowLayout">申請日期：</DIV>
			<asp:button id="Go" style="Z-INDEX: 111; LEFT: 464px; POSITION: absolute; TOP: 77px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button>
            <asp:TextBox ID="DItemName1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" MaxLength="35" ReadOnly="True" Style="z-index: 112;
                left: 342px; position: absolute; top: 34px" Width="235px">ItemName1</asp:TextBox>
            <asp:TextBox ID="DItemName2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" MaxLength="35" ReadOnly="True" Style="z-index: 113;
                left: 583px; position: absolute; top: 33px" Width="235px">ItemName2</asp:TextBox>
            <asp:TextBox ID="DItemName3" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" MaxLength="35" ReadOnly="True" Style="z-index: 114;
                left: 583px; position: absolute; top: 59px" Width="235px">ItemName3</asp:TextBox>
            <asp:TextBox ID="DItemCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" MaxLength="35" ReadOnly="True" Style="z-index: 115;
                left: 583px; position: absolute; top: 4px" Width="235px">ItemCode</asp:TextBox>

            <asp:dropdownlist id="DKeepData" style="Z-INDEX: 102; LEFT: 380px; POSITION: absolute; TOP: 77px" runat="server"
				Height="26px" Width="79px" ForeColor="Blue" BackColor="Yellow">
                <asp:ListItem Selected="True" Value="0">未封存</asp:ListItem>
                <asp:ListItem Value="1">已封存</asp:ListItem>
            </asp:DropDownList>
        </form>
	</body>
</HTML>
