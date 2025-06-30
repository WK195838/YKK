<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentDelivery_OP.aspx.vb" Inherits="DevelopmentDelivery_OP" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>工程</title>
</head>
<body>
    <form id="form1" runat="server">
<!-- *************************************************************************************************************** -->
<!-- #[開發履歷] -->
        <asp:HyperLink ID="LHistory" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 20px; position: absolute; top: 15px; text-align: center" Target="_blank"
            Width="30px">@</asp:HyperLink>
<!-- #[工程1] -->
        <asp:Image ID="DOP1Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 12px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!-- -->
<!--    工程範圍(DOP1Step) -->
        <asp:TextBox ID="DOP1Step" runat="server" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 105px; position: absolute;
            top: 19px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger" BackColor="Transparent">@@@</asp:TextBox>
<!--    工程名稱(DOP1Name) -->
        <asp:TextBox ID="DOP1Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 19px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP1Count) -->
        <asp:TextBox ID="DOP1Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 19px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP1Yotei) -->
        <asp:TextBox ID="DOP1Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 49px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP1Jisseki) -->
        <asp:TextBox ID="DOP1Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 49px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP1Detail) -->
        <asp:HyperLink ID="LOP1Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 106px; position: absolute; top: 77px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP1Loading) -->
        <asp:HyperLink ID="LOP1Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 140px; position: absolute; top: 77px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP1BDate) -->
        <asp:TextBox ID="DOP1BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 77px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP1BTime) -->
        <asp:TextBox ID="DOP1BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 335px; position: absolute;
            top: 77px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP1ADate) -->
        <asp:TextBox ID="DOP1ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 397px; position: absolute;
            top: 77px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP1ATime) -->
        <asp:TextBox ID="DOP1ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 558px; position: absolute;
            top: 77px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程2] -->
        <asp:Image ID="DOP2Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 120px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP2Step) -->
        <asp:TextBox ID="DOP2Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 105px; position: absolute;
            top: 129px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP2Name) -->
        <asp:TextBox ID="DOP2Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 129px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP2Count) -->
        <asp:TextBox ID="DOP2Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 129px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP2Yotei) -->
        <asp:TextBox ID="DOP2Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 157px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP2Jisseki) -->
        <asp:TextBox ID="DOP2Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 157px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP2Detail) -->
        <asp:HyperLink ID="LOP2Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 105px; position: absolute; top: 185px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP2Loading) -->
        <asp:HyperLink ID="LOP2Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 139px; position: absolute; top: 185px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP2BDate) -->
        <asp:TextBox ID="DOP2BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 185px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP2BTime) -->
        <asp:TextBox ID="DOP2BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 335px; position: absolute;
            top: 185px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP2ADate) -->
        <asp:TextBox ID="DOP2ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 397px; position: absolute;
            top: 185px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP2ATime) -->
        <asp:TextBox ID="DOP2ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 558px; position: absolute;
            top: 185px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程3] -->
        <asp:Image ID="DOP3Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 228px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP3Step) -->
        <asp:TextBox ID="DOP3Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 105px; position: absolute;
            top: 235px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP3Name) -->
        <asp:TextBox ID="DOP3Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 235px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP3Count) -->
        <asp:TextBox ID="DOP3Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 235px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP3Yotei) -->
        <asp:TextBox ID="DOP3Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 265px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP3Jisseki) -->
        <asp:TextBox ID="DOP3Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 265px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP3Detail) -->
        <asp:HyperLink ID="LOP3Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 105px; position: absolute; top: 293px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP3Loading) -->
        <asp:HyperLink ID="LOP3Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 139px; position: absolute; top: 293px; text-align: center" Target="_blank"
            Width="30px" EnableTheming="True">負</asp:HyperLink>
<!--    預定完成日(DOP3BDate) -->
        <asp:TextBox ID="DOP3BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 293px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP3BTime) -->
        <asp:TextBox ID="DOP3BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 335px; position: absolute;
            top: 293px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP3ADate) -->
        <asp:TextBox ID="DOP3ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 397px; position: absolute;
            top: 293px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP3ATime) -->
        <asp:TextBox ID="DOP3ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 558px; position: absolute;
            top: 293px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程4] -->
        <asp:Image ID="DOP4Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 336px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP4Step) -->
        <asp:TextBox ID="DOP4Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 105px; position: absolute;
            top: 344px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP4Name) -->
        <asp:TextBox ID="DOP4Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 343px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP4Count) -->
        <asp:TextBox ID="DOP4Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 343px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP4Yotei) -->
        <asp:TextBox ID="DOP4Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 373px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP4Jisseki) -->
        <asp:TextBox ID="DOP4Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 373px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP4Detail) -->
        <asp:HyperLink ID="LOP4Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 105px; position: absolute; top: 401px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP4Loading) -->
        <asp:HyperLink ID="LOP4Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 139px; position: absolute; top: 401px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP4BDate) -->
        <asp:TextBox ID="DOP4BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 401px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP4BTime) -->
        <asp:TextBox ID="DOP4BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 335px; position: absolute;
            top: 401px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP4ADate) -->
        <asp:TextBox ID="DOP4ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 397px; position: absolute;
            top: 401px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP4ATime) -->
        <asp:TextBox ID="DOP4ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 558px; position: absolute;
            top: 401px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程5] -->
        <asp:Image ID="DOP5Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 444px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP5Step) -->
        <asp:TextBox ID="DOP5Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 105px; position: absolute;
            top: 452px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP5Name) -->
        <asp:TextBox ID="DOP5Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 451px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP5Count) -->
        <asp:TextBox ID="DOP5Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 451px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP5Yotei) -->
        <asp:TextBox ID="DOP5Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 481px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP5Jisseki) -->
        <asp:TextBox ID="DOP5Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 481px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP5Detail) -->
        <asp:HyperLink ID="LOP5Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 105px; position: absolute; top: 509px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP5Loading) -->
        <asp:HyperLink ID="LOP5Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 139px; position: absolute; top: 509px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP5BDate) -->
        <asp:TextBox ID="DOP5BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
            top: 509px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP5BTime) -->
        <asp:TextBox ID="DOP5BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 335px; position: absolute;
            top: 509px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP5ADate) -->
        <asp:TextBox ID="DOP5ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 397px; position: absolute;
            top: 509px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP5ATime) -->
        <asp:TextBox ID="DOP5ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 558px; position: absolute;
            top: 509px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程6] -->
        <asp:Image ID="DOP6Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 552px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP6Step) -->
        <asp:TextBox ID="DOP6Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 106px; position: absolute;
            top: 559px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP6Name) -->
        <asp:TextBox ID="DOP6Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 177px; position: absolute;
            top: 559px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP6Count) -->
        <asp:TextBox ID="DOP6Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 559px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP6Yotei) -->
        <asp:TextBox ID="DOP6Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 177px; position: absolute;
            top: 589px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP6Jisseki) -->
        <asp:TextBox ID="DOP6Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 399px; position: absolute;
            top: 589px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP6Detail) -->
        <asp:HyperLink ID="LOP6Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 106px; position: absolute; top: 617px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP6Loading) -->
        <asp:HyperLink ID="LOP6Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 140px; position: absolute; top: 617px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP6BDate) -->
        <asp:TextBox ID="DOP6BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 177px; position: absolute;
            top: 617px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP6BTime) -->
        <asp:TextBox ID="DOP6BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 336px; position: absolute;
            top: 617px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP6ADate) -->
        <asp:TextBox ID="DOP6ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 617px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP6ATime) -->
        <asp:TextBox ID="DOP6ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 559px; position: absolute;
            top: 617px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程7] -->
        <asp:Image ID="DOP7Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 660px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP7Step) -->
        <asp:TextBox ID="DOP7Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 106px; position: absolute;
            top: 670px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP7Name) -->
        <asp:TextBox ID="DOP7Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 177px; position: absolute;
            top: 667px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP7Count) -->
        <asp:TextBox ID="DOP7Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 667px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP7Yotei) -->
        <asp:TextBox ID="DOP7Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 177px; position: absolute;
            top: 698px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP7Jisseki) -->
        <asp:TextBox ID="DOP7Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 399px; position: absolute;
            top: 698px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP7Detail) -->
        <asp:HyperLink ID="LOP7Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 106px; position: absolute; top: 725px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP7Loading) -->
        <asp:HyperLink ID="LOP7Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 140px; position: absolute; top: 725px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP7BDate) -->
        <asp:TextBox ID="DOP7BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 177px; position: absolute;
            top: 725px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP7BTime) -->
        <asp:TextBox ID="DOP7BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 336px; position: absolute;
            top: 725px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP7ADate) -->
        <asp:TextBox ID="DOP7ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 398px; position: absolute;
            top: 725px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP7ATime) -->
        <asp:TextBox ID="DOP7ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 559px; position: absolute;
            top: 725px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程8] -->
        <asp:Image ID="DOP8Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 768px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP8Step) -->
        <asp:TextBox ID="DOP8Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 107px; position: absolute;
            top: 775px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP8Name) -->
        <asp:TextBox ID="DOP8Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 178px; position: absolute;
            top: 775px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP8Count) -->
        <asp:TextBox ID="DOP8Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 775px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP8Yotei) -->
        <asp:TextBox ID="DOP8Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 178px; position: absolute;
            top: 804px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP8Jisseki) -->
        <asp:TextBox ID="DOP8Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 400px; position: absolute;
            top: 804px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP8Detail) -->
        <asp:HyperLink ID="LOP8Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 107px; position: absolute; top: 833px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP8Loading) -->
        <asp:HyperLink ID="LOP8Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 141px; position: absolute; top: 833px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP8BDate) -->
        <asp:TextBox ID="DOP8BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 178px; position: absolute;
            top: 833px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP8BTime) -->
        <asp:TextBox ID="DOP8BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 337px; position: absolute;
            top: 833px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP8ADate) -->
        <asp:TextBox ID="DOP8ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 399px; position: absolute;
            top: 833px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP8ATime) -->
        <asp:TextBox ID="DOP8ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 560px; position: absolute;
            top: 833px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!-- *************************************************************************************************************** -->
<!-- #[工程9] -->
        <asp:Image ID="DOP9Image" runat="server" Style="z-index: 1; left: 100px; position: absolute;
            top: 876px" ImageUrl="Images\DevelopmentDelivery_01.jpg" />
<!--    工程範圍(DOP9Step) -->
        <asp:TextBox ID="DOP9Step" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 107px; position: absolute;
            top: 883px; text-align: center" Width="62px" Font-Bold="True" Font-Size="Larger">@@@</asp:TextBox>
<!--    工程名稱(DOP9Name) -->
        <asp:TextBox ID="DOP9Name" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 178px; position: absolute;
            top: 883px; text-align: center" Width="437px" Font-Bold="True" Font-Size="Larger">委託確立</asp:TextBox>
<!--    工程件數(DOP9Count) -->
        <asp:TextBox ID="DOP9Count" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 576px; position: absolute;
            top: 883px; text-align: center" Width="40px" Font-Bold="True" Font-Size="Larger">(2)</asp:TextBox>
<!--    預定(DOP9Yotei) -->
        <asp:TextBox ID="DOP9Yotei" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 178px; position: absolute;
            top: 912px; text-align: center" Width="215px" Font-Size="Larger">預定</asp:TextBox>
<!--    實績(DOP9Jisseki) -->
        <asp:TextBox ID="DOP9Jisseki" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 400px; position: absolute;
            top: 912px; text-align: center" Width="215px" Font-Size="Larger">實績</asp:TextBox>
<!--    細項(LOP9Detail) -->
        <asp:HyperLink ID="LOP9Detail" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 107px; position: absolute; top: 941px; text-align: center" Target="_blank"
            Width="30px">細</asp:HyperLink>
<!--    負荷(LOP9Loading) -->
        <asp:HyperLink ID="LOP9Loading" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 141px; position: absolute; top: 941px; text-align: center" Target="_blank"
            Width="30px">負</asp:HyperLink>
<!--    預定完成日(DOP9BDate) -->
        <asp:TextBox ID="DOP9BDate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 178px; position: absolute;
            top: 941px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    預定分數(DOP9BTime) -->
        <asp:TextBox ID="DOP9BTime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 337px; position: absolute;
            top: 941px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
<!--    實績完成日(DOP9ADate) -->
        <asp:TextBox ID="DOP9ADate" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 399px; position: absolute;
            top: 941px; text-align: center" Width="154px" Font-Size="Larger">2012/12/12 12:12</asp:TextBox>
<!--    實績分數(DOP9ATime) -->
        <asp:TextBox ID="DOP9ATime" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="Blue"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 560px; position: absolute;
            top: 941px; text-align: right" Width="56px" Font-Size="Larger">9999</asp:TextBox>
    </form>
</body>
</html>
