<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintCustomer.aspx.vb" Inherits="MaintCustomer" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div id="DIV1" runat="server">
    <table >
    <tr>
    <td>
        &nbsp;<img src="images/CustomerTable.jpg" /></td>
    </tr>
    
    <tr><td align="right">
        &nbsp;<asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" /><asp:Button ID="Button2" runat="server" Text="儲存" /></td></tr>
    </table>
    </div>
        <asp:TextBox ID="DSMTPFirstTime" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 100; left: 136px; position: absolute; top: 461px" Width="210px"></asp:TextBox>
        <asp:TextBox ID="DSMTPPeriod" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 101; left: 596px; position: absolute; top: 428px" Width="90px"></asp:TextBox>
        <asp:TextBox ID="DSMTPLastTime" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 102; left: 473px; position: absolute; top: 461px" Width="210px"></asp:TextBox>
        <asp:TextBox ID="DPDFLastTime" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 103; left: 475px; position: absolute; top: 361px" Width="210px"></asp:TextBox>
        <asp:TextBox ID="DPDFPeriod" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 104; left: 596px; position: absolute; top: 330px" Width="90px"></asp:TextBox>
        <asp:TextBox ID="DPDFFirstTime" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 105; left: 136px; position: absolute; top: 362px" Width="210px"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" Style="z-index: 106;
            left: 296px; position: absolute; top: 197px" Width="225px" SkinID="Txt_Yellow" MaxLength="50"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="Textbox1" runat="server" Style="z-index: 107;
            left: 296px; position: absolute; top: 263px" Width="225px" SkinID="Txt_Green" MaxLength="50"></asp:TextBox>
        <asp:TextBox ID="DMailToPosition2" runat="server" SkinID="Txt_Yellow" Style="z-index: 108;
            left: 536px; position: absolute; top: 197px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSalesCode" runat="server" SkinID="Txt_Green" Style="z-index: 122;
            left: 536px; position: absolute; top: 263px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DCustCode" runat="server" SkinID="Txt_Green" Style="z-index: 110;
            left: 296px; position: absolute; top: 131px" Width="164px" MaxLength="7"></asp:TextBox>
        <asp:TextBox ID="DMailToPosition1" runat="server" SkinID="Txt_Yellow" Style="z-index: 111;
            left: 536px; position: absolute; top: 163px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" Style="z-index: 115;
            left: 296px; position: absolute; top: 163px" Width="225px" SkinID="Txt_Yellow" MaxLength="50"></asp:TextBox>
        <asp:TextBox ID="DCustName" runat="server" SkinID="Txt_Green" Style="z-index: 113;
            left: 136px; position: absolute; top: 131px" MaxLength="50"></asp:TextBox>
        <asp:TextBox ID="DSalesMan" runat="server" SkinID="Txt_Green" Style="z-index: 114;
            left: 136px; position: absolute; top: 263px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DMailToName1" runat="server" SkinID="Txt_Yellow" Style="z-index: 115;
            left: 136px; position: absolute; top: 163px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DMailToName2" runat="server" SkinID="Txt_Yellow" Style="z-index: 116;
            left: 136px; position: absolute; top: 197px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DUniqueID" runat="server" ReadOnly="True" SkinID="Txt_Lightgray"
            Style="z-index: 117; left: 477px; position: absolute; top: 97px" Width="210px"></asp:TextBox>
        <asp:DropDownList ID="DSMTPSend" runat="server" SkinID="Ddl_Green" Style="z-index: 118;
            left: 136px; position: absolute; top: 427px" Width="110px">
            <asp:ListItem Value="1">是</asp:ListItem>
            <asp:ListItem Value="0">否</asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DSMTPResend" runat="server" SkinID="Ddl_Green" Style="z-index: 118;
            left: 368px; position: absolute; top: 427px" Width="110px">
            <asp:ListItem Value="1">是</asp:ListItem>
            <asp:ListItem Value="0" Selected="True">否</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DPDFCreate" runat="server" SkinID="Ddl_Green" Style="z-index: 119;
            left: 136px; position: absolute; top: 329px" Width="110px">
            <asp:ListItem Value="1">是</asp:ListItem>
            <asp:ListItem Value="0">否</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DPDFRecreate" runat="server" SkinID="Ddl_Green" Style="z-index: 119;
            left: 368px; position: absolute; top: 329px" Width="110px">
            <asp:ListItem Value="1">是</asp:ListItem>
            <asp:ListItem Value="0" Selected="True">否</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DMailCCList" runat="server" SkinID="Ddl_Green" Style="z-index: 120;
            left: 136px; position: absolute; top: 230px" Width="150px">
            <asp:ListItem Value="0">無</asp:ListItem>
            <asp:ListItem Value="1">有</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DActive" runat="server" SkinID="Ddl_Green" Style="z-index: 121;
            left: 136px; position: absolute; top: 97px" Width="150px">
            <asp:ListItem Value="1">啟用</asp:ListItem>
            <asp:ListItem Value="0">停用</asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DIntCust" runat="server" SkinID="Ddl_Green" Style="z-index: 121;
            left: 469px; position: absolute; top: 131px" Width="150px">
            <asp:ListItem Selected="True" Value="1">國內</asp:ListItem>
            <asp:ListItem Value="0">國外</asp:ListItem>
        </asp:DropDownList>
        &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="客戶名稱需輸入" ControlToValidate="DCustName" Display="None"></asp:RequiredFieldValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
            ShowSummary="False" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="DCustCode"
            Display="None" ErrorMessage="客戶編號需輸入"></asp:RequiredFieldValidator>
        &nbsp;&nbsp;
        <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ControlToValidate="DSalesMan"
            Display="None" ErrorMessage="業務名稱需輸入"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ControlToValidate="Textbox1"
            Display="None" ErrorMessage="業務郵件需輸入"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ControlToValidate="DSalesCode"
            Display="None" ErrorMessage="業務編號需輸入"></asp:RequiredFieldValidator>
    </form>
</body>
</html>
