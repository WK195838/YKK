<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SP_ShoppingListInf.aspx.vb" Inherits="SP_ShoppingListInf" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Shopping List Inf.</title>
	    <script language="javascript" src="ForProject_SP.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
        Height="18px" ReadOnly="True" Style="z-index: 126; left: 16px; position: absolute;
        top: 16px; text-align: left" Width="80px">SP Name</asp:TextBox>
    <asp:TextBox ID="DKSPName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
        Height="18px" Style="z-index: 126; left: 96px; position: absolute;
        top: 16px; text-align: left" Width="200px" MaxLength="30"></asp:TextBox>

    <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
        Height="18px" ReadOnly="True" Style="z-index: 126; left:304px; position: absolute;
        top: 16px; text-align: left" Width="48px">Search</asp:TextBox>
    <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
        Height="18px" Style="z-index: 126; left: 352px; position: absolute;
        top: 16px; text-align: left" Width="208px" MaxLength="50"></asp:TextBox>

    <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 576px;
            position: absolute; top: 16px" Text="Go" Width="45px" />

<asp:TextBox style="Z-INDEX: 126; LEFT: 632px; POSITION: absolute; TOP: 16px; TEXT-ALIGN: left" id="DDescription" runat="server"  ForeColor="black" Height="18px" Width="536px" BorderStyle="Groove" BackColor="white" MaxLength="7">說明</asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Data                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    <asp:GridView ID="GridView1" runat="server" WIDTH="1000px" AutoGenerateColumns="False"  Style="z-index: 103; left: 16px; position: absolute; top: 48px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
        <Columns>
            <asp:BoundField DataField="SPCode"  />
            <asp:BoundField DataField="SPName"  />
            <asp:BoundField DataField="SPTime"  />
            <asp:BoundField DataField="SPNo"  />
            <asp:BoundField DataField="Status" />
            <asp:BoundField DataField="LastAcessTime" />
            <asp:BoundField DataField="ChangeFinal"  >
                <ItemStyle HorizontalAlign="Center" />
            </asp:BoundField >
            
            <asp:HyperLinkField DataNavigateUrlFields="ChangeUrl" DataNavigateUrlFormatString="{0}" 
            DataTextField="Change" HeaderText="Link" Target="_blank"  >
                <HeaderStyle HorizontalAlign="Center" />
                <ItemStyle HorizontalAlign="Center" />
            </asp:HyperLinkField>       

            <asp:BoundField DataField="UserList" />

        </Columns>
        <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
        <FooterStyle BackColor="#CCCCCC" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
        <AlternatingRowStyle BackColor="#CCCCCC" />
    </asp:GridView>
    </div>
    </form>
</body>
</html>
