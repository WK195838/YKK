<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EDIFlow.aspx.vb" Inherits="EDIFlow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>EDI System 2011 Ver2.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body >
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/Flow151110.png" style="z-index: 1; left: -4px; position: absolute;top: 34px; width: 953px; height: 608px;"/>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  公告執行跑馬燈                                                                      ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
		<asp:hyperlink id="LBoard" style="Z-INDEX: 102; LEFT: 840px; POSITION: absolute; TOP: 13px" runat="server"
					Width="103px" Height="16px" Font-Size="12pt" NavigateUrl="Document/納期說明文件.xlsx" Target="_blank" Enabled="true">納期統一說明</asp:hyperlink>
<asp:TextBox style="Z-INDEX: 126; LEFT: 620px; POSITION: absolute; TOP: 10px; TEXT-ALIGN: left" id="DBoard" runat="server" MaxLength="35" Width="328px" Height="18px" ForeColor="Red" BorderStyle="Groove" >2023/1/1起實施 RCA切換 (目的:RC收支改善)</asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BActionLog" runat="server" style="z-index: 1; left: 2px; position: absolute;top: 0px;" ImageUrl="iMages/ActionLog.jpg" Height="38px" Width="38px" />
        <asp:ImageButton ID="BCheckBuyerPrice" runat="server" style="z-index: 1; left: 50px; position: absolute;top: 0px;" ImageUrl="iMages/CheckBuyerPrice.png" Height="37px" Width="70px" />
        <asp:ImageButton ID="BProgress" runat="server" style="z-index: 1; left: 130px; position: absolute;top: 0px;" ImageUrl="iMages/ProgressBar.png" Height="38px" Width="70px" />
        <asp:ImageButton ID="BPackList" runat="server" style="z-index: 1; left: 210px; position: absolute;top: 14px;" ImageUrl="~/iMages/packinglist.jpg" Height="24px" Width="70px" />
        <asp:ImageButton ID="BSCIJobList" runat="server" style="z-index: 1; left: 290px; position: absolute;top: 8px;" ImageUrl="~/iMages/SCIJobList.png" Height="30px" Width="60px" />
        <asp:ImageButton ID="BDSJobList" runat="server" style="z-index: 1; left: 290px; position: absolute;top: 8px;" ImageUrl="~/iMages/DSJobList.png" Height="30px" Width="60px" />
        <asp:ImageButton ID="BQMIJobList" runat="server" style="z-index: 1; left: 290px; position: absolute;top: 8px;" ImageUrl="~/iMages/QMIJobList.jpg" Height="30px" Width="60px" />
        <asp:ImageButton ID="BLULUWIP" runat="server" style="z-index: 1; left: 290px; position: absolute;top: 8px;" ImageUrl="~/iMages/LULUWIP.jpg" Height="30px" Width="60px" />
        <asp:ImageButton ID="BEOES" runat="server" style="z-index: 1; left: 530px; position: absolute;top: 8px;" ImageUrl="~/iMages/EOES.jpg" Height="30px" Width="60px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:CheckBox ID="AtOffice2010" runat="server" style="z-index: 174; left: 192px; position: absolute; top: 131px" Font-Size="9pt" Text="Office2010" Width="81px" Checked="True" Enabled="false" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BCustomer" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 82px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 142px; position: absolute;top: 82px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" />
        <asp:ImageButton ID="BExcel" runat="server" style="z-index: 1; left: 165px; position: absolute;top: 302px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />
        <asp:ImageButton ID="BImport" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 522px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BDataCheck" runat="server" style="z-index: 1; left: 554px; position: absolute;top: 82px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BWaves" runat="server" style="z-index: 1; left: 917px; position: absolute;top: 82px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BSalesPrice" runat="server" style="z-index: 1; left: 917px; position: absolute;top: 485px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查資料-狀態                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:HyperLink ID="StsCustomer" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 167px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsExcel" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 377px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsImport" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 572px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>

        <asp:HyperLink ID="StsMakePONO" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 214px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckPONO" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 258px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckGRPC" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 302px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckCompanyCode" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 346px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckKeepCode" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 390px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckColorCode" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 434px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckItemCode" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 478px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckDuplicateData" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 522px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsCheckPOStructure" runat="server" style="z-index: 1; left: 520px; position: absolute;top: 566px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>

        <asp:HyperLink ID="StsWaves" runat="server" style="z-index: 1; left: 917px; position: absolute;top: 132px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsSalesPrice" runat="server" style="z-index: 1; left: 917px; position: absolute;top: 535px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Check Information                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:HyperLink ID="Inf_PriceList" runat="server" style="z-index: 1; left: 732px; position: absolute;top: 598px;" Height="20px" Width="212px" ForeColor="Navy" NavigateUrl="InfPriceList.aspx" Target="_blank">Sales Price Information</asp:HyperLink>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <marquee id="ProExcel" runat="server" style="z-index: 1; left: 7px; position: absolute;top: 399px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProImport" runat="server" style="z-index: 1; left: 8px; position: absolute;top: 605px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProDataCheck" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 185px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProWaves" runat="server" style="z-index: 1; left: 732px; position: absolute;top: 185px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProSalesPrice" runat="server" style="z-index: 1; left: 732px; position: absolute;top: 576px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  客戶選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:DropDownList ID="DCustomerBuyer" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 9px; position: absolute; top: 163px" Width="172px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DLogID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DGRBuyer" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
