<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MailAddress.aspx.vb" Inherits="MailAddress" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>郵件帳號查詢</title>
</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV style="DISPLAY: inline; Z-INDEX: 103; LEFT: 8px; WIDTH: 56px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">公司：</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 106; LEFT: 616px; POSITION: absolute; TOP: 8px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:dropdownlist id="DDivision" style="Z-INDEX: 105; LEFT: 264px; POSITION: absolute; TOP: 8px"
					runat="server" Height="40px" Width="128px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist>
				<DIV style="DISPLAY: inline; Z-INDEX: 104; LEFT: 208px; WIDTH: 56px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">部門：</DIV>
				<asp:dropdownlist id="DDepoName" style="Z-INDEX: 101; LEFT: 64px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="128px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist>
				<DIV style="DISPLAY: inline; Z-INDEX: 107; LEFT: 408px; WIDTH: 65px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">
                    姓名：</DIV>
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                    BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
                    Style="z-index: 114; left: 10px; position: absolute; top: 44px" Width="500px">
                    <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                    <Columns>
                        <asp:BoundField DataField="DepoNameDesc" HeaderText="公司" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DivisionDesc" HeaderText="部門" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NameDesc" HeaderText="姓名" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                <asp:HyperLinkField DataNavigateUrlFields="MailURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="I_MailAddress" HeaderText="Internet Mail Address" Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                </asp:HyperLinkField>
                    </Columns>
                    <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:GridView>
                <asp:TextBox ID="DFirstName" runat="server" BackColor="Yellow" BorderStyle="Groove"
                    ForeColor="Blue" Height="20px" Style="z-index: 103; left: 477px; position: absolute;
                    top: 8px" Width="82px"></asp:TextBox>
                <asp:Button ID="BGo" runat="server" Height="25px" Style="z-index: 104; left: 567px;
                    position: absolute; top: 8px" Text="Go" Width="45px" />
            </FONT></form>
	</body>
</html>
