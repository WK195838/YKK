<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Main.aspx.vb" Inherits="Main" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>SSES System Ver 1.0</title>
	<script language="javascript" src="ForProject.js"></script>
</head>

<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="Images/3s.png" style="z-index: 1; left: -4px; position: absolute;top: 4px; width: 1070px; height: 470px;"/>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 174px; position: absolute;top: 20px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" />
        <asp:ImageButton ID="BJGSFun" runat="server" style="z-index: 1; left: 174px; position: absolute;top: 187px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        <asp:ImageButton ID="BRISFun" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 376px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        <asp:ImageButton ID="BJGS" runat="server" style="z-index: 1; left: 536px; position: absolute;top: 62px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        <asp:ImageButton ID="BRIS" runat="server" style="z-index: 1; left: 401px; position: absolute;top: 193px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        <asp:ImageButton ID="BBIS" runat="server" style="z-index: 1; left: 848px; position: absolute;top: 62px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        <asp:ImageButton ID="BDGS" runat="server" style="z-index: 1; left: 848px; position: absolute;top: 254px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        <asp:ImageButton ID="BRBS" runat="server" style="z-index: 1; left: 848px; position: absolute;top: 396px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" />
        &nbsp;
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <marquee id="ProJGSFun" runat="server" style="z-index: 1; left: 7px; position: absolute;top: 234px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProRISFun" runat="server" style="z-index: 1; left: 8px; position: absolute;top: 424px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProJGS" runat="server" style="z-index: 1; left: 400px; position: absolute;top: 177px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProRIS" runat="server" style="z-index: 1; left: 266px; position: absolute;top: 361px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProBIS" runat="server" style="z-index: 1; left: 870px; position: absolute;top: 159px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProDGS" runat="server" style="z-index: 1; left: 870px; position: absolute;top: 326px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProRBS" runat="server" style="z-index: 1; left: 870px; position: absolute;top: 449px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  機能選擇                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:DropDownList ID="DJGSFun" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 6px; position: absolute; top: 202px" Width="170px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DRISFun" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 8px; position: absolute; top: 392px" Width="170px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
        &nbsp;

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查資料-狀態                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:HyperLink ID="StsJGSFun" runat="server" style="z-index: 1; left: 174px; position: absolute;top: 187px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>
        <asp:HyperLink ID="StsRISFun" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 376px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>
        <asp:HyperLink ID="StsJGS" runat="server" style="z-index: 1; left: 536px; position: absolute;top: 62px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>
        <asp:HyperLink ID="StsRIS" runat="server" style="z-index: 1; left: 401px; position: absolute;top: 193px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>
        <asp:HyperLink ID="StsBIS" runat="server" style="z-index: 1; left: 848px; position: absolute;top: 62px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>
        <asp:HyperLink ID="StsDGS" runat="server" style="z-index: 1; left: 848px; position: absolute;top: 254px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>
        <asp:HyperLink ID="StsRBS" runat="server" style="z-index: 1; left: 848px; position: absolute;top: 396px; width: 40px; height: 40px;" ImageUrl="iMages/OK.png"/>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="xFun" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 13px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="xFunID" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 23px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="xFileName" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 33px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="xFilePath" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

    </div>
    </form>
</body>
</html>
