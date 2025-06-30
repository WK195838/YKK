<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GESSMenu.aspx.vb" Inherits="GESSMenu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>GESS System 2020 Ver1.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/GESS_MENU.png" style="z-index: 1; left: 40px; position: absolute;top: 0px; width: 1000px; height: 550px;"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- START -->
        <asp:ImageButton ID="BStart" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
            left: 166px; width: 40px; position: absolute; top: 24px; height: 40px" />
        <!-- CONVERT / MODIFY -->
        <asp:ImageButton ID="BConvert" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
            left: 360px; width: 40px; position: absolute; top: 224px; height: 40px" />
        <asp:ImageButton ID="BModify" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
            left: 824px; width: 40px; position: absolute; top: 208px; height: 40px" />
        <!-- MAINTCONVERT / INQHISTORY -->
        <asp:ImageButton ID="BMaintConvert" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
            left: 616px; width: 40px; position: absolute; top: 392px; height: 40px" />
        <asp:ImageButton ID="BInqHistory" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
            left: 828px; width: 40px; position: absolute; top: 392px; height: 40px" />
        <!-- M.OTHER -->
       <asp:ImageButton ID="BOther" runat="server" style="z-index: 1; left: 256px; position: absolute;top: 24px;" ImageUrl="iMages/GESS_Btn_Other.png" Height="40px" Width="90px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  下拉式選單                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:DropDownList ID="DStart" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 48px; position: absolute; top: 32px" Width="120px">
            <asp:ListItem>開始使用</asp:ListItem>
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath2" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath3" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath4" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath5" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    </div>
    </form>
</body>

</html>