<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemRequest_02.aspx.vb" Inherits="ItemRequest_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Item Request Sheet</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/ItemRequest_02.jpg" style="z-index: 1; left: 6px; position: absolute;top: 7px" />

        <asp:Button ID="BFindItem" runat="server" CausesValidation="False" Height="22px" Style="z-index: 131;
            left: 118px; position: absolute; top: 10px" Text="搜尋參考Item" Width="100px" />
        <asp:Button ID="BSave" runat="server" Height="22px" Style="z-index: 104;
            left: 223px; position: absolute; top: 10px" Text="儲存" Width="100px" />
        <asp:Button ID="BDelete" runat="server" Height="22px" Style="z-index: 104;
            left: 326px; position: absolute; top: 10px" Text="刪除" Width="100px" />
        <asp:Button ID="BClose" runat="server" Height="22px" Style="z-index: 104;
            left: 429px; position: absolute; top: 10px" Text="關閉" Width="100px" />
            
        <asp:TextBox ID="DRCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 65px; text-align: left" Width="289px" MaxLength="7"></asp:TextBox>
        <asp:TextBox ID="DRItemName1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 91px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRItemName2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 117px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRItemName3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 143px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRSize" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 169px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DRChain" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 195px; text-align: left" Width="289px" MaxLength="6"></asp:TextBox>
        <asp:TextBox ID="DRClass" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 221px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DRTape" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 247px; text-align: left" Width="289px" MaxLength="5"></asp:TextBox>
        <asp:TextBox ID="DRSlider1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 273px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DRFinish1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 299px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DRSlider2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 325px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DRFinish2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 351px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DRSRequest1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 377px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 403px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 429px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 455px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 481px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 507px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRFamily" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 533px; text-align: left" Width="289px" MaxLength="4"></asp:TextBox>
        <asp:TextBox ID="DRST1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 157px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 198px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 239px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 280px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 321px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 362px;
            position: absolute; top: 559px; text-align: left" Width="43px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DItemName1" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Red"
            Height="20px" Style="z-index: 126; left: 410px; position: absolute;
            top: 91px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DCode" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Red"
            Height="20px" MaxLength="35" Style="z-index: 126; left: 410px; position: absolute;
            top: 65px; text-align: left" Width="289px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DID" runat="server" BackColor="Silver" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="35" Style="z-index: 126; left: 554px; position: absolute;
            top: 12px; text-align: left" Width="66px"></asp:TextBox>
        <asp:TextBox ID="DItemName2" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Red"
            Height="20px" Style="z-index: 126; left: 410px; position: absolute;
            top: 117px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DItemName3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 143px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DSize" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 169px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DChain" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 195px; text-align: left" Width="289px" MaxLength="6"></asp:TextBox>
        <asp:TextBox ID="DClass" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 221px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DTape" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 247px; text-align: left" Width="289px" MaxLength="5"></asp:TextBox>
        <asp:TextBox ID="DSlider1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 273px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DFinish1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 299px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DSlider2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 325px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DFinish2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 351px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DSRequest1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 377px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 403px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 429px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 455px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 481px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest6" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 507px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DFamily" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 533px; text-align: left" Width="289px" MaxLength="4"></asp:TextBox>
        <asp:TextBox ID="DST1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 451px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 492px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 533px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 574px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST6" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 615px;
            position: absolute; top: 559px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST7" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 656px;
            position: absolute; top: 559px; text-align: left" Width="43px" MaxLength="1"></asp:TextBox>
        <asp:DropDownList ID="DSts" runat="server" BackColor="Yellow" ForeColor="Blue" Style="z-index: 112;
            left: 627px; position: absolute; top: 11px" Width="73px">
            <asp:ListItem Value="0">Ｏ</asp:ListItem>
            <asp:ListItem Value="1">Ｘ</asp:ListItem>
            <asp:ListItem Value="2">Ｗ</asp:ListItem>
        </asp:DropDownList>
        </div>
    </form>
</body>
</html>
