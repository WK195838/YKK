<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TNFBYFlow.aspx.vb" Inherits="TNFBYFlow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>FAS Buyer 2013 Ver1.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/BYMain_NIKE.png" style="z-index: 1; left: 0px; position: absolute;top: 34px; width: 996px; height: 420px;"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BStatusList" runat="server" style="z-index: 1; left: 2px; position: absolute;top: 0px;" ImageUrl="iMages/BStatus.png" Height="35px" Width="86px" Visible="False" />
        <asp:ImageButton ID="BCheckSum" runat="server" style="z-index: 1; left: 89px; position: absolute;top: 0px;" ImageUrl="iMages/summary-icon.gif" Height="35px" Width="86px" Visible="False" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 162px; position: absolute;top: 77px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" Visible="False" />
        <asp:ImageButton ID="BBYFMTChange" runat="server" style="z-index: 1; left: 212px; position: absolute;top: 77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BBYImport" runat="server" style="z-index: 1; left: 165px; position: absolute;top: 314px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />
        <!-- 2列  -->
        <asp:ImageButton ID="BBYDataCheck" runat="server" style="z-index: 1; left: 552px; position: absolute;top:77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BBYReport" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 313px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 3列  -->
        <asp:ImageButton ID="BBYConvert" runat="server" style="z-index: 1; left: 912px; position: absolute;top:77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BBYFASReport" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 313px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查資料-狀態                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:HyperLink ID="StsBYFMTChange" runat="server" style="z-index: 1; left: 212px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsBYImport" runat="server" style="z-index: 1; left: 165px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 2列  -->
        <asp:HyperLink ID="StsBYDataCheck" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsBYReport" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 3列  -->
        <asp:HyperLink ID="StsBYConvert" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsBYFASReport" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <marquee id="ProBYFMTChange" runat="server" style="z-index: 1; left: 10px; position: absolute;top: 198px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProBYImport" runat="server" style="z-index: 1; left: 10px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 2列  -->
        <marquee id="ProBYDataCheck" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 198px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProBYReport" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 3列  -->
        <marquee id="ProBYConvert" runat="server" style="z-index: 1; left: 730px; position: absolute;top: 198px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProBYFASReport" runat="server" style="z-index: 1; left: 730px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  客戶選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- BY FCT格式轉換  -->
        <asp:CheckBox ID="AtGoImport" runat="server" style="z-index: 174; left: 258px; position: absolute; top: 82px" Font-Size="9pt" Text="至準備BYFCT資料" Width="130px" Visible="False" />
        <asp:CheckBox ID="AtGoReport" runat="server" style="z-index: 174; left: 258px; position: absolute; top: 107px" Font-Size="9pt" Text="至報表列出" Width="130px" Visible="False" />
        <!-- 準備 BY FCT資料  -->
        <asp:CheckBox ID="AtNotFirst" runat="server" style="z-index: 174; left: 218px; position: absolute; top: 345px" Font-Size="9pt" Text="更新BYFCT" Width="115px" />
        <!-- ITEM / COLOR CONVERT -->
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
        <asp:TextBox ID="DLastUniqueID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath2" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- ReportData  -->
        <asp:TextBox ID="DReportFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DReportFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
