<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ActionHistory.aspx.vb" Inherits="ActionHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>系統履歷</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
				 
        <asp:DropDownList ID="DSts" runat="server" BackColor="Yellow"
            Font-Size="10pt" ForeColor="Blue" Height="26px" Style="z-index: 103; left: 225px;
            position: absolute; top: 8px" Width="90px">
        </asp:DropDownList>
 
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 8px; position: absolute; top: 40px" Width="1200px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" Wrap="True" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="Err1" HeaderText="訂單" Target="_blank">
                </asp:HyperLinkField>
                <asp:HyperLinkField DataNavigateUrlFields="URL1" DataNavigateUrlFormatString="{0}"
                    DataTextField="Err2" HeaderText="Waves" Target="_blank">
                </asp:HyperLinkField>
                                
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
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 110; left: 625px; position: absolute; top: 7px" Text="Go" Width="40px" />
        <asp:TextBox ID="DLine" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 103; left: 536px; position: absolute;
            top: 8px" Width="80px"></asp:TextBox>
        <asp:TextBox ID="DCustBuyer" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 103; left: 9px; position: absolute;
            top: 8px" Width="93px">H4530-000013</asp:TextBox>
        <asp:DropDownList ID="DAction" runat="server" BackColor="Yellow"
            Font-Size="10pt" ForeColor="Blue" Height="26px" Style="z-index: 103; left: 318px;
            position: absolute; top: 8px" Width="216px">
            </asp:DropDownList>
        &nbsp;
        <asp:TextBox ID="DLogID" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 103; left: 110px; position: absolute; top: 8px"
            Width="107px">20121224115736</asp:TextBox>
    </div>
    </form>
</body>
</html>
