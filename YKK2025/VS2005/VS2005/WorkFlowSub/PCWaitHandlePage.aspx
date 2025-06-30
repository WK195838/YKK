<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PCWaitHandlePage.aspx.vb" Inherits="PCWaitHandlePage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>待處理通知</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Image ID="MessageTop" runat="server" Height="65px" ImageUrl="Images\MailTop.jpg"
            Style="left: 2px; position: absolute; top: 0px" Width="334px" />
        <asp:Image ID="MessageBottom" runat="server" Height="3px" ImageUrl="Images\MailBottom.jpg"
            Style="left: 2px; position: absolute; top: 146px" Width="327px" />
        <asp:Label ID="Label1" runat="server" Height="20px" Style="z-index: 103; left: 74px;
            position: absolute; top: 92px" Width="48px">急　件</asp:Label>
        <asp:Label ID="DHigh" runat="server" Font-Size="14pt" Font-Underline="True" ForeColor="#C00000"
            Height="20px" Style="z-index: 103; left: 134px; position: absolute; top: 91px; text-align: right;"
            Width="48px">(9999)</asp:Label>
        <asp:Label ID="Label2" runat="server" ForeColor="Red" Height="20px" Style="z-index: 103;
            left: 2px; position: absolute; top: 60px; text-align: center;" Width="317px" BorderStyle="None">＊有新的待核定電子文件＊</asp:Label>
        <asp:Label ID="Label3" runat="server" Height="20px" Style="z-index: 103; left: 74px;
            position: absolute; top: 121px" Width="48px">一般件</asp:Label>
        <asp:Label ID="DLow" runat="server" Font-Size="14pt" Font-Underline="True" ForeColor="#C00000"
            Height="20px" Style="z-index: 103; left: 134px; position: absolute; top: 120px; text-align: right;"
            Width="48px">(9999)</asp:Label>
        <asp:HyperLink ID="DLabel1" runat="server" Font-Size="10pt" Font-Underline="False"
            NavigateUrl="http://10.245.1.10/Portal/工作流程/待處理/tabid/90/Default.aspx" Style="left: 226px;
            position: absolute; top: 133px; text-align: right" Target="_blank" Width="88px">立即處理？</asp:HyperLink>
    
    </div>
    </form>
</body>
</html>
