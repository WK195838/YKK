<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CategoriPicker.aspx.vb" Inherits="CategoriPicker" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DDKey" runat="server" SkinID="Txt_Yellow" Width="60%"></asp:TextBox><asp:Button
            ID="Button1" runat="server" Text="...." />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DDKey"
            Display="None" ErrorMessage="不允許空白"></asp:RequiredFieldValidator>
        &nbsp;<asp:CompareValidator ID="CompareValidator1" runat="server" ControlToValidate="DDKey"
            Display="None" ErrorMessage="不允許999" Operator="NotEqual" ValueToCompare="999"></asp:CompareValidator>&nbsp;
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
            ShowSummary="False" />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            SkinID="GV" Width="80%">
            <Columns>
                <asp:TemplateField HeaderText="點選">
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" CommandArgument='<%# Eval("DKey") %>' CausesValidation="False" CommandName="Pick">選取</asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="類別">
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("DKey") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="名稱">
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Eval("Data") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
