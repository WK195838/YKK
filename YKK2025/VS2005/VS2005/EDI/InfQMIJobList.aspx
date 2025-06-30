<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfQMIJobList.aspx.vb" Inherits="InfQMIJobList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>QMI Job List</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>

<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/JobList.png" style="z-index: 1; left: 350px; position: absolute;top: 34px; width: 230px; height: 590px;"/>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  各按鈕                                                                              ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BJob" runat="server" style="z-index: 1; left: 541px; position: absolute;top: 154px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BExcel" runat="server" style="z-index: 1; left: 541px; position: absolute;top: 352px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />
        <asp:ImageButton ID="BRun" runat="server" style="z-index: 1; left: 541px; position: absolute;top: 523px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 541px; position: absolute;top: 73px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" />
    
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  檢查資料-狀態                                                                       ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:HyperLink ID="StsJob" runat="server" style="z-index: 1; left: 601px; position: absolute;top: 154px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsExcel" runat="server" style="z-index: 1; left: 601px; position: absolute;top: 352px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsRun" runat="server" style="z-index: 1; left: 601px; position: absolute;top: 523px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  執行跑馬燈                                                                       ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <marquee id="ProExcel" runat="server" style="z-index: 1; left: 357px; position: absolute;top: 395px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProRun" runat="server" style="z-index: 1; left: 357px; position: absolute;top: 395px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  工作選擇                                                                       ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:DropDownList ID="DJob" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 360px; position: absolute; top: 170px" Width="172px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  隱藏欄位                                                                       ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFileName1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
