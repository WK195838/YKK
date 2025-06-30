<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ColorInquiry.aspx.vb" Inherits="ColorInquiry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>ITEM-查詢</title>
</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 7px; WIDTH: 44px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                Season</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 122; LEFT: 506px; POSITION: absolute; TOP: 41px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px"></asp:imagebutton><asp:textbox id="DSeason" style="Z-INDEX: 116; LEFT: 54px; POSITION: absolute; TOP: 39px" runat="server"
				Height="20px" Width="140px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
			<asp:button id="Go" style="Z-INDEX: 110; LEFT: 457px; POSITION: absolute; TOP: 39px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button>
            <asp:TextBox ID="DColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" Style="z-index: 111;
                left: 302px; position: absolute; top: 39px" Width="140px"></asp:TextBox>
            <DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 208px; WIDTH: 92px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                Color1 / 2</div>
            <DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 7px; WIDTH: 44px; COLOR: #0000ff; POSITION: absolute; TOP: 7px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                Buyer</div>
            <asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow"
                Font-Size="10pt" ForeColor="Blue" Height="26px" Style="z-index: 103; left: 54px;
                position: absolute; top: 8px" Width="145px">
            </asp:DropDownList>
                
                
                
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 72px" Width="800px">
                <Columns>
                    <asp:BoundField DataField="BuyerDesc" HeaderText="Buyer" ReadOnly="True" />
                    <asp:BoundField DataField="Season" HeaderText="Season" ReadOnly="True" />
                    <asp:BoundField DataField="Color1" HeaderText="Color-1" ReadOnly="True" />
                    <asp:BoundField DataField="Color2" HeaderText="Color-2" ReadOnly="True" />
                    <asp:BoundField DataField="Green" HeaderText="Green" ReadOnly="True" />
                    <asp:BoundField DataField="YColor" HeaderText="YKK-COLOR" ReadOnly="True" />
                    <asp:BoundField DataField="YColorName" HeaderText="YKK-Name" ReadOnly="True" />
                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
        </form>
	</body>
</html>
