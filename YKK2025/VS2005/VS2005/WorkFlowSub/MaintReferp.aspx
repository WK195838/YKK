<%@ Page Language="VB"  AutoEventWireup="false" CodeFile="MaintReferp.aspx.vb" Inherits="MaintReferp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
 
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DCat" runat="server" BackColor="Yellow" Width="171px"></asp:TextBox>
        &nbsp;
        <input id="BSelect" type="button" value="..." runat="server" />&nbsp;&nbsp;<asp:Button ID="BSerch" runat="server" Text="搜尋　" EnableViewState="False" />
        <asp:Button ID="BADD" runat="server" Text="新項目　" />&nbsp;&nbsp;&nbsp;<br />
        &nbsp;&nbsp;
        <br />
        <asp:TextBox ID="DDKey" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="16px" Style="z-index: 112; left: 11px; position: absolute; top: 41px"
            Width="170px">DKey</asp:TextBox>
        <br />
        類別：&nbsp;
        &nbsp;&nbsp;
        <asp:TextBox ID="DName" runat="server" Style="z-index: 102; left: 176px; position: absolute;
            top: 80px" Width="228px" BackColor="White" BorderStyle="None"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderStyle="None" Style="z-index: 104;
            left: 63px; position: absolute; top: 79px" Width="88px"></asp:TextBox>
        <br />
        
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 10px; position: absolute; top: 100px" Width="400px" AllowPaging="True" PageSize="20">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="BUpd" runat="server" CommandArgument='<%# eval("Unique_ID") %>' Text="變更" CommandName="Upd"   />
                        <asp:Button ID="BDel" runat="server" CommandArgument='<%# eval("Unique_ID") %>' Text="刪除" CommandName="Del" OnClientClick="return confirm('是否要刪除?')"   />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Cat" HeaderText="類別" />
                <asp:BoundField DataField="Dkey" HeaderText="項目 " />
                <asp:BoundField DataField="Data" HeaderText="設定/名稱" />
            </Columns>
                <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                <RowStyle BackColor="White" ForeColor="Blue" />
                <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
                <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
                <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>

