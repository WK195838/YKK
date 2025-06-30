<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DASW_BatchAutoApprove.aspx.vb" Inherits="DASW_BatchAutoApprove" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>批次自動簽核處理</title>
</head>
	<body MS_POSITIONING="GridLayout">
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DTextBox1" runat="server" AutoPostBack="True" BackColor="#E0E0E0"
            BorderStyle="Groove" Font-Bold="False" Font-Size="14pt" ForeColor="Black" Height="28px"
            MaxLength="7" Style="z-index: 126; left: 5px; position: absolute; top: 7px; text-align: left"
            Width="461px">[HRWCL] 加班自動簽核處理</asp:TextBox>
        <asp:Label ID="Label1" runat="server" Text="FormNo = " Style="z-index: 126; left: 7px; position: absolute; top: 44px; text-align: left" Height="24px" Width="98px" Visible="False"></asp:Label>
        <asp:Label ID="Label2" runat="server" Height="24px" Style="z-index: 126; left: 7px;
            position: absolute; top: 69px; text-align: left" Text="FormSno = " Width="97px" Visible="False"></asp:Label>
        <asp:Label ID="Label3" runat="server" Height="24px" Style="z-index: 126; left: 7px;
            position: absolute; top: 121px; text-align: left" Text="簽核者ID = " Width="97px" Visible="False"></asp:Label>
        <asp:Label ID="Label4" runat="server" Height="24px" Style="z-index: 126; left: 7px;
            position: absolute; top: 95px; text-align: left" Text="Step = " Width="97px" Visible="False"></asp:Label>
        <asp:Label ID="Label5" runat="server" Height="24px" Style="z-index: 126; left: 7px;
            position: absolute; top: 148px; text-align: left" Text="行事曆群組 = " Width="97px" Visible="False"></asp:Label>
        <asp:Label ID="DMsg" runat="server" ForeColor="Red" Height="24px" Style="z-index: 126;
            left: 7px; position: absolute; top: 212px; text-align: left" Width="460px" Visible="False"></asp:Label>
        <asp:TextBox ID="DFormNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 105; left: 108px;
            position: absolute; top: 42px" Width="88px" Visible="False">001052</asp:TextBox>
        <asp:TextBox ID="DFormSno" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 105; left: 108px; position: absolute;
            top: 68px" Width="88px" Visible="False"></asp:TextBox>
        <asp:TextBox ID="DStep" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 105; left: 108px; position: absolute; top: 94px"
            Width="88px" Visible="False"></asp:TextBox>
        <asp:TextBox ID="DDecideID" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 105; left: 108px; position: absolute;
            top: 121px" Width="88px" Visible="False"></asp:TextBox>
        <asp:TextBox ID="DcalendarGr" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 105; left: 108px; position: absolute;
            top: 148px" Width="88px" Visible="False"></asp:TextBox>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="46px"
            Style="z-index: 102; left: 10px; position: absolute; top: 242px" Text="Go" Width="109px" Visible="False" />
        <asp:Label ID="Label6" runat="server" Height="24px" Style="z-index: 126; left: 7px;
            position: absolute; top: 176px; text-align: left" Text="指定下一簽核者 = " Width="133px" Visible="False"></asp:Label>
        <asp:TextBox ID="DAllocateID" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 105; left: 142px; position: absolute;
            top: 176px" Width="88px" Visible="False"></asp:TextBox>
        <asp:DropDownList ID="DNextGate" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 127; left: 256px; position: absolute; top: 176px" Width="240px" AppendDataBoundItems="True">
        </asp:DropDownList>
    </form>
</body>
</html>
