<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_VDPErrorList.aspx.vb" Inherits="InfF_VDPErrorList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>VDP Inf. Error List</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  篩選欄位                                                                            ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">Buyer：</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 156px; position: absolute; top: 8px">Year：</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 296px; position: absolute; top: 8px">Month：</asp:Label>

        <asp:CheckBox ID="AtSeason" runat="server" style="z-index: 174; left: 13px; position: absolute; top: 35px" Font-Size="9pt" Text="Season" Width="72px" Checked="True" Enabled="False" />
        <asp:CheckBox ID="AtPLM" runat="server" style="z-index: 174; left: 90px; position: absolute; top: 35px" Font-Size="9pt" Text="PLM" Width="72px" Checked="True" Enabled="False" />
        <asp:CheckBox ID="AtBuyMonth" runat="server" style="z-index: 174; left: 167px; position: absolute; top: 35px" Font-Size="9pt" Text="BuyMonth" Width="89px" Checked="True" />
        <asp:CheckBox ID="AtStyle" runat="server" style="z-index: 174; left: 261px; position: absolute; top: 35px" Font-Size="9pt" Text="Style" Width="72px" Checked="True" />
        <asp:CheckBox ID="AtADIDASColor" runat="server" style="z-index: 174; left: 338px; position: absolute; top: 35px" Font-Size="9pt" Text="Color" Width="109px" />
        
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 64px; position: absolute; top: 8px">
        </asp:TextBox>

        <asp:DropDownList id="DYY" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 204px; position: absolute; top: 8px">
            <asp:ListItem>2013</asp:ListItem>
        </asp:DropDownList>

        <asp:DropDownList id="DMM" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 353px; position: absolute; top: 8px">
            <asp:ListItem>1</asp:ListItem>
        </asp:DropDownList>
        
        <asp:TextBox ID="DBlank" runat="server" BackColor="White" BorderStyle="None" ForeColor="#C00000" Height="16px" Style="z-index: 103; left: 486px; position: absolute; top: 10px" Width="179px">實際Order資料時點=當日7:00
        </asp:TextBox>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  按鈕                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 482px; position: absolute; top: 34px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>
    	
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  ITEM GridView                                                                   ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" Style="z-index: 103; left: 15px; position: absolute; top: 62px" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt"  GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Completed" HeaderText="Status"  />
                <asp:BoundField DataField="CustWavesCode" HeaderText="成衣廠"  />
                <asp:BoundField DataField="OrderNo" HeaderText="Order"  />
                <asp:BoundField DataField="OrderDate" HeaderText="Confirm Date"  />

                <asp:BoundField DataField="V_Season" HeaderText="Season Code"  />
                <asp:BoundField DataField="V_Year" HeaderText="Season Year"  />
                <asp:BoundField DataField="Season" HeaderText="Season"  />
                <asp:BoundField DataField="V_BuyMonth" HeaderText="Buy Month"  />
                <asp:BoundField DataField="V_PLM" HeaderText="PLM"  />
                <asp:BoundField DataField="V_Style" HeaderText="Style"  />
                <asp:BoundField DataField="V_TapeColor" HeaderText="Color"  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

   </div>
   </form>
</body>
</html>
