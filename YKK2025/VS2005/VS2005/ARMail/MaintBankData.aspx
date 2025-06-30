<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintBankData.aspx.vb" Inherits="MaintBankData" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/MaintBank.jpg" style="z-index: 1; left: 13px; position: absolute;
            top: 17px" />
        <asp:TextBox ID="DBankName" runat="server" Style="z-index: 100; left: 143px; position: absolute;
            top: 194px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:TextBox ID="DBankACNo" runat="server" Style="z-index: 101; left: 143px; position: absolute;
            top: 259px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:TextBox ID="DBankAddress" runat="server" Style="z-index: 102; left: 143px; position: absolute;
            top: 228px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:TextBox ID="DCustName" runat="server" Style="z-index: 103; left: 143px; position: absolute;
            top: 127px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:TextBox ID="DSalesCode" runat="server" Style="z-index: 104; left: 143px; position: absolute;
            top: 160px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:TextBox ID="DSwift" runat="server" Style="z-index: 105; left: 143px; position: absolute;
            top: 325px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:TextBox ID="DCustCode" runat="server" Style="z-index: 110; left: 143px; position: absolute;
            top: 94px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Style="z-index: 107; left: 637px; position: absolute;
            top: 365px" Text="確定" Width="40px" />
        &nbsp;
        <asp:TextBox ID="DBankACName" runat="server" Style="z-index: 108; left: 143px; position: absolute;
            top: 293px" SkinID="Txt_Green" Width="500px"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DCustCode"
            Display="None" ErrorMessage="請輸入Customer Code"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="DCustName"
            Display="None" ErrorMessage="請輸入Customer Name"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="DSalesCode"
            Display="None" ErrorMessage="請輸入Sales Code"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="DBankName"
            Display="None" ErrorMessage="請輸入Bank Name"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="DBankAddress"
            Display="None" ErrorMessage="請輸入Bank Address"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ControlToValidate="DBankACNo"
            Display="None" ErrorMessage="請輸入Bank ACNO"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ControlToValidate="DBankACName"
            Display="None" ErrorMessage="請輸入Bank ACName"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ControlToValidate="DSwift"
            Display="None" ErrorMessage="請輸入Swift"></asp:RequiredFieldValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
            ShowSummary="False" />
        <asp:HiddenField ID="DID" runat="server" />
    
    </div>
    </form>
</body>
</html>
