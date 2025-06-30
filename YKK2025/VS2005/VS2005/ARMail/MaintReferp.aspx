<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintReferp.aspx.vb" Inherits="MaintReferp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        類別：<asp:TextBox ID="TextBox1" runat="server" ReadOnly="True" SkinID="txt_Yellow"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="...." CausesValidation="False" />&nbsp;
        <br />
        項目：<asp:TextBox ID="TextBox2" runat="server" SkinID="txt_Yellow"></asp:TextBox>
        <asp:Button ID="Button6" runat="server" Text="搜尋" />
        <asp:Button ID="Button3" runat="server" Text="新項目" CausesValidation="False" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
            Display="None" ErrorMessage="不允許空白"></asp:RequiredFieldValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
            ShowSummary="False" />
        <asp:HiddenField ID="HiddenField1" runat="server" />
        <br />
        <br />
        <asp:Label ID="Label1" runat="server"></asp:Label><br />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" SkinID="GV" Width="80%" AllowPaging="True" PageSize="50">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="Button5" runat="server" CommandArgument='<%# Eval("Unique_ID") & "," & Eval("Level") %>'
                            CommandName="Upt" Text="變更" />
                        <asp:Button ID="Button4" runat="server" CommandArgument='<%# Eval("Unique_ID") & "," & Eval("Level") %>'
                            CommandName="Del" Text="刪除" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Cat" HeaderText="類別" />
                <asp:BoundField DataField="DKey" HeaderText="項目" />
                <asp:BoundField DataField="Data" HeaderText="設定/名稱" />
            </Columns>
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
