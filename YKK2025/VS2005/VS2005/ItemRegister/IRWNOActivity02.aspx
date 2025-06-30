<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNOActivity02.aspx.vb" Inherits="IRWNOActivity02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>ITEM No Activity-02</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 103; left: 4px; position: absolute; top: 8px" Width="24px" />

        <asp:DropDownList ID="DYYMM" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="24px" Style="z-index: 112; left: 40px; position: absolute; top: 8px"
            Width="168px">
            <asp:ListItem Selected="True" Value="BUYERITEM">ITEM</asp:ListItem>
            <asp:ListItem Value="BUYERCOLOR-TAPE">COLOR(TAPE)</asp:ListItem>
            <asp:ListItem Value="BUYERCOLOR-PULLER">COLOR(PULLER)</asp:ListItem>
            <asp:ListItem Value="BUYERPRICE">PRICE</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DStatus" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="24px" Style="z-index: 112; left: 208px; position: absolute; top: 8px"
            Width="168px">
            <asp:ListItem Selected="True" Value="ALL">ALL</asp:ListItem>
            <asp:ListItem Value="NG">NG</asp:ListItem>
            <asp:ListItem Value="YES">YES</asp:ListItem>
        </asp:DropDownList>

<asp:button style="Z-INDEX: 103; LEFT: 384px; POSITION: absolute; TOP: 8px" id="BGo" runat="server" Width="48px" Height="24px" Font-Size="Small" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"  Width="1500px" Style="z-index: 103; left: 0px; position: absolute; top: 32px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="200" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Status" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 

                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:BoundField DataField="DepName" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="Name" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="xDate" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="xCompletedTime" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="ForUse" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 

                <asp:BoundField DataField="Code" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="Itemname1" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="Itemname2" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="Itemname3" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 

                <asp:BoundField DataField="Customer" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="Buyer" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                
                <asp:BoundField DataField="CCode" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
                <asp:BoundField DataField="WNo" HeaderText="" ItemStyle-HorizontalAlign="Left"/> 
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