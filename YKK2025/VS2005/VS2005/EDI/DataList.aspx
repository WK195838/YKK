<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DataList.aspx.vb" Inherits="DataList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>EDI資料一覽</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>

        <asp:DropDownList ID="DCustomer" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="28px" Style="z-index: 102; left: 11px; position: absolute;
            top: 8px" Width="140px">
        </asp:DropDownList>
        
        <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 154px; position: absolute; top: 7px"
            Width="112px"></asp:TextBox>

        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 108; left: 276px; position: absolute; top: 7px" Text="Go" Width="40px" />

            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="Black"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 41px">
                <Columns>
                    <asp:BoundField DataField="Buyer" HeaderText="客戶" />
                    <asp:BoundField DataField="DataDate" HeaderText="日期" />
                    <asp:BoundField DataField="DataTime" HeaderText="時間" />
                    
                 <asp:HyperLinkField DataNavigateUrlFields="EDI" DataNavigateUrlFormatString="{0}"
                    DataTextField="EDIDesc" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                 <asp:HyperLinkField DataNavigateUrlFields="EDIAll" DataNavigateUrlFormatString="{0}"
                    DataTextField="EDIAllDesc" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                 <asp:HyperLinkField DataNavigateUrlFields="EDIOri" DataNavigateUrlFormatString="{0}"
                    DataTextField="EDIOriDesc" HeaderText="" Target="_blank">
                </asp:HyperLinkField>
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
    
    </div>
    </form>
</body>
</html>
