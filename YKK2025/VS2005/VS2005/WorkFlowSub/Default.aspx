<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head runat="server">
    <title>委託單</title>
</head>

<body>
    <form id="form1" runat="server">
    <div>
        FormNo:<asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem  Value="001101">案件委託</asp:ListItem>
        </asp:DropDownList>
        FormSNo<asp:TextBox ID="TextBox2" runat="server" Width="95px">0</asp:TextBox>Step<asp:TextBox ID="TextBox3" runat="server"  Width="43px">1</asp:TextBox>
        SeqNo<asp:TextBox ID="TextBox4" runat="server" Width="66px">0</asp:TextBox>
        ApplyID<asp:TextBox ID="TextBox5" runat="server"  Width="67px">it003</asp:TextBox>
        UserID<asp:TextBox ID="TextBox6" runat="server"  Width="67px">it003</asp:TextBox>
        <br />
        <asp:Button ID="CaseReviewSheet_01" runat="server" Text="案件委託_01" />
        <asp:Button ID="CaseReviewSheet_02" runat="server" Text="案件委託_02" />
        <asp:Button ID="CaseReviewSheet_H" runat="server" Text="履歷" />&nbsp;<br />
        &nbsp;
        <br />
        &nbsp;
        <br />
        &nbsp; &nbsp;<br />
        &nbsp; &nbsp;<br />
        <br />
        </div>
        
    </form>
</body>
</html>
