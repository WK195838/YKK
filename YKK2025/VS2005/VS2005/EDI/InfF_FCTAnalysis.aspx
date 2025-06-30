<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_FCTAnalysis.aspx.vb" Inherits="InfF_FCTAnalysis" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>FCT-FCT Customer</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">Buyer：</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 424px; position: absolute; top: 8px">Season：</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 148px; position: absolute; top: 8px">Year：</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 281px; position: absolute; top: 8px">Month：</asp:Label>

        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>

        <asp:DropDownList id="DSeason" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 486px; position: absolute; top: 8px">
            <asp:ListItem>SS14</asp:ListItem>
        </asp:DropDownList>

        <asp:DropDownList id="DYY" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 196px; position: absolute; top: 8px" AutoPostBack="True">
            <asp:ListItem>2013</asp:ListItem>
        </asp:DropDownList>

        <asp:DropDownList id="DMM" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 338px; position: absolute; top: 8px" AutoPostBack="True">
            <asp:ListItem>1</asp:ListItem>
        </asp:DropDownList>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 含 New-FCT -->
        <asp:CheckBox ID="CNewFCT" runat="server" style="z-index: 174; left: 570px; position: absolute; top: 11px" Font-Size="10pt" Text="含 New-FCT" Width="100px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 675px; position: absolute; top: 10px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="CSTGridView" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 40px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="Customer" HeaderText="成衣廠" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="Season" HeaderText="季"  />
                
                <asp:BoundField DataField="Qty6"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio6" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty5"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio5" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty4"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio4" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty3"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio3" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty2"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio2" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Qty1"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio1" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
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
