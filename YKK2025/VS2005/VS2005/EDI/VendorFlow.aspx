<%@ Page Language="VB" AutoEventWireup="false" CodeFile="VendorFlow.aspx.vb" Inherits="VendorFlow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Vendor FC </title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/VendorFlow.png" style="z-index: 1; left: -6px; position: absolute;top: 0px; width: 945px; height: 410px;"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:ImageButton ID="BSalesMan" runat="server" style="z-index: 1; left: 200px; position: absolute;top: 43px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 250px; position: absolute;top: 43px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" />
        <asp:ImageButton ID="BUpdateFC" runat="server" style="z-index: 1; left: 200px; position: absolute;top: 284px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />
        <!-- 2列  -->
        <asp:ImageButton ID="BAppendFC" runat="server" style="z-index: 1; left: 562px; position: absolute;top:43px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BConfirm" runat="server" style="z-index: 1; left: 562px; position: absolute;top: 284px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 3列  -->
        <asp:ImageButton ID="BFinal" runat="server" style="z-index: 1; left: 925px; position: absolute;top:43px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 4列  -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查資料-狀態                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:HyperLink ID="StsSalesMan" runat="server" style="z-index: 1; left: 142px; position: absolute;top: 93px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsUpdateFC" runat="server" style="z-index: 1; left: 165px; position: absolute;top: 330px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 2列  -->
        <asp:HyperLink ID="StsAppendFC" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 93px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsConfirm" runat="server" style="z-index: 1; left: 525px; position: absolute;top: 330px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 3列  -->
        <asp:HyperLink ID="StsFinal" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 93px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 4列  -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <marquee id="ProUpdateFC" runat="server" style="z-index: 1; left: 10px; position: absolute;top: 406px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 2列  -->
        <marquee id="ProAppendFC" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 164px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProConfirm" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 406px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 3列  -->
        <marquee id="ProFinal" runat="server" style="z-index: 1; left: 730px; position: absolute;top: 164px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 4列  -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  客戶選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:DropDownList ID="DSalesMan" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 9px; position: absolute; top: 149px" Width="172px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:CheckBox ID="AtAppendFC" runat="server" style="z-index: 174; left: 202px; position: absolute; top: 111px" Font-Size="9pt" Text="新FC INF." Width="100px" />
        <asp:CheckBox ID="AtConfirm" runat="server" style="z-index: 174; left: 202px; position: absolute; top: 131px" Font-Size="9pt" Text="確認FC INF." Width="100px" />
        <asp:CheckBox ID="AtFinal" runat="server" style="z-index: 174; left: 202px; position: absolute; top: 151px" Font-Size="9pt" Text="移送管理人" Width="100px" />
        <!-- Office2010 -->
        <asp:CheckBox ID="AtOffice2010" runat="server" style="z-index: 174; left: 202px; position: absolute; top: 91px" Font-Size="9pt" Text="Office2010" Width="82px" Checked="True" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="DLogID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DGRBuyer" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFunList" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFileName1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- ReportData  -->
    </div>
    </form>
</body>
</html>
