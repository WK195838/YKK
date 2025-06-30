<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPSActionHistory.aspx.vb" Inherits="SPSActionHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>系統履歷</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
				 
        <asp:TextBox ID="DCustBuyer" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 103; left: 9px; position: absolute;
            top: 8px" Width="93px">H4530-000013</asp:TextBox>
        <asp:TextBox ID="DLogID" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 103; left: 110px; position: absolute; top: 8px"
            Width="107px">20121224115736</asp:TextBox>
 
        <asp:Button ID="BError" runat="server"   Height="24px"
            Style="z-index: 110; left: 232px; position: absolute; top: 8px" Text="Error" Width="80px" />

        <asp:Button ID="BAll" runat="server"  Height="24px"
            Style="z-index: 110; left: 320px; position: absolute; top: 8px" Text="All" Width="80px" />

        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 8px; position: absolute; top: 40px" >
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" Wrap="True" />
            <Columns>
                                
                <asp:BoundField DataField="ActionDesc" HeaderText="功能" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="RunTimeDesc" HeaderText="時間" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="ErrorDesc" HeaderText="狀態" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="D1" HeaderText="說明-1" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="D2" HeaderText="說明-2" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="D3" HeaderText="說明-3" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="D4" HeaderText="說明-4" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="D5" HeaderText="說明-5" ReadOnly="True">
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
    </div>
    </form>
</body>

</html>
