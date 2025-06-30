<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintBank.aspx.vb" Inherits="MaintBank" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="TextBox1" runat="server" SkinID="Txt_Yellow" ></asp:TextBox>
        <asp:Button ID="Button1" runat="server"  Text="查詢" />
        <asp:Button ID="Button4" runat="server"  Text="新增" /><br />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" SkinID="GV" >
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="Button3" runat="server" CommandName="Upt" Text="變更" CommandArgument='<%# Eval("CustCode") %>' />
                        <asp:Button ID="Button2" runat="server" CommandName="Del" Text="刪除" CommandArgument='<%# Eval("CustCode") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Customer">
                    
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("CustCode") & "-" & Eval("CustName") %>'></asp:Label>
                        
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="SalesCode" HeaderText="Sales Code" />
                <asp:BoundField DataField="BankName" HeaderText="Bank Name" />
                <asp:BoundField DataField="BankAddress" HeaderText="Bank Address" />
                <asp:BoundField DataField="BankACNO" HeaderText="Bank ACNo" />
                <asp:BoundField DataField="BankACName" HeaderText="Bank ACName" />
                <asp:BoundField DataField="SWIFT" HeaderText="Swift" />
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:OAOPEN_Con %>" SelectCommand="SELECT * FROM [M_Bank]  where 1=1 ORDER BY [CustCode], [CustName]"></asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
