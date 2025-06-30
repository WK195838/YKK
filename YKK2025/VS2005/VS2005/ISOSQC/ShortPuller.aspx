<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ShortPuller.aspx.vb" Inherits="ShortPuller" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-Short Puller</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/ShortPuller.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 48px; position: absolute;
            top: 56px; text-align: left" Width="80px">Puller Code</asp:TextBox>
        <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 128px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
        &nbsp; &nbsp;
        &nbsp;&nbsp;
        <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 432px; position: absolute;
            top: 56px; text-align: left" Width="80px">Search</asp:TextBox>
        <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="50" Style="z-index: 126; left: 512px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 632px;
            position: absolute; top: 56px" Text="Go" Width="45px" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 48px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>

                <asp:BoundField DataField="Puller" HeaderText=""  />      
                <asp:BoundField DataField="Short" HeaderText=""  />          
                <asp:BoundField DataField="BUYER" HeaderText=""  />      
                
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>      
        <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 240px; position: absolute;
            top: 56px; text-align: left" Width="80px">Buyer</asp:TextBox>
        <asp:TextBox ID="DKBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 320px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
