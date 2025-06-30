<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_FCTACTAnalysis.aspx.vb" Inherits="InfF_FCTACTAnalysis" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>ACT-FCT Customer</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 17px; position: absolute; top: 10px">Buyer：</asp:Label>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 71px; position: absolute; top: 10px"></asp:TextBox>

        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 156px; position: absolute;
            top: 10px">Version：</asp:Label>
        <asp:DropDownList id="DVersion" runat="server" Width="129px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 221px; position: absolute; top: 10px">
            <asp:ListItem>20140830121212</asp:ListItem>
        </asp:DropDownList>
        
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 357px; position: absolute; top: 10px">Season：</asp:Label>
        <asp:DropDownList id="DSeason" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 419px; position: absolute; top: 10px">
            <asp:ListItem>SS14</asp:ListItem>
        </asp:DropDownList>

        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 507px; position: absolute; top: 9px">Type：</asp:Label>
        <asp:DropDownList id="DLevel" runat="server" Width="124px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 560px; position: absolute; top: 7px">
            <asp:ListItem Selected="True">ITEM</asp:ListItem>
            <asp:ListItem>ITEM+COLOR</asp:ListItem>
        </asp:DropDownList>

        <asp:TextBox ID="DBlank" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="#C00000" Height="16px" Style="z-index: 103; left: 18px; position: absolute;
            top: 39px" Width="290px">實際Order資料時點=當日7:00 (受注日基準)</asp:TextBox>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  Action (CheckBox)                                                                   ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- VDP -->
        <!-- VDP -->

            
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 694px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="CSTGridView" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 15px; position: absolute; top: 62px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="Customer" HeaderText="成衣廠" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="Season" HeaderText="季"  />
                
                <asp:BoundField DataField="AQty6"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FQty6"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio6" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="AQty5"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FQty5"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio5" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="AQty4"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FQty4"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio4" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="AQty3"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FQty3"  > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="Ratio3" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
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
