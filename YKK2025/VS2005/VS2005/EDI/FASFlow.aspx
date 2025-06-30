<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FASFlow.aspx.vb" Inherits="FASFlow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>FAS 2013 Ver1.0</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/FAS_Main_1.png" style="z-index: 1; left: -6px; position: absolute;top: 34px; width: 940px; height: 645px;" id="IMG1"/>
        <img src="iMages/FAS_Main_2.png" style="z-index: 1; left: 931px; position: absolute;top: 34px; width: 360px;" height="645"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- 固定(上方) -->
        <asp:ImageButton ID="BActionLog" runat="server" style="z-index: 1; left: 2px; position: absolute;top: 0px;" ImageUrl="iMages/ActionLog.jpg" Height="35px" Width="35px" />
        <asp:ImageButton ID="BBuyerFCT" runat="server" style="z-index: 1; left: 38px; position: absolute;top: 0px;" ImageUrl="iMages/BuyerForecast.jpg" Height="35px" Width="75px" />
        <asp:ImageButton ID="BLSInput" runat="server" style="z-index: 1; left: 38px; position: absolute;top: 0px;" ImageUrl="iMages/BLSInput.jpg" Height="35px" Width="75px" />
        <asp:ImageButton ID="BItemCheck" runat="server" style="z-index: 1; left: 114px; position: absolute;top: 0px;" ImageUrl="iMages/ItemCheck.jpg" Height="35px" Width="84px" />
<!-- PAGE-1(上方) -->
        <asp:ImageButton ID="BLEFT" runat="server" style="z-index: 1; left: 199px; position: absolute;top: 0px;" ImageUrl="iMages/LEFT.jpg" Height="35px" Width="54px" />
        <asp:ImageButton ID="BPLANR" runat="server" style="z-index: 1; left: 254px; position: absolute;top: 0px;" ImageUrl="iMages/PLAN_R.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BFCT2ACTR" runat="server" style="z-index: 1; left: 339px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2ACT_R.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BFCT2ACTMMR" runat="server" style="z-index: 1; left: 424px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2ACTMM_R.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BFCT2ACTVDPR" runat="server" style="z-index: 1; left: 509px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2ACTVDP_R.jpg" Height="35px" Width="84px" />
<!-- 材料分析不使用 換成KPI -->
        <asp:ImageButton ID="BMAT2SEA" runat="server" style="z-index: 1; left: 594px; position: absolute;top: 0px;" ImageUrl="iMages/MAT2SEA.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BMAT2MON" runat="server" style="z-index: 1; left: 679px; position: absolute;top: 0px;" ImageUrl="iMages/MAT2MON.jpg" Height="35px" Width="84px" />

        <asp:ImageButton ID="BKPI" runat="server" style="z-index: 1; left: 594px; position: absolute;top: 0px;" ImageUrl="iMages/eKPI.png" Height="35px" Width="84px" />


        <asp:ImageButton ID="BFCT2ORD" runat="server" style="z-index: 1; left: 764px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2ORD.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BFCT2FCT" runat="server" style="z-index: 1; left: 849px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2FCT.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BFCT2ACT" runat="server" style="z-index: 1; left: 934px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2ACT.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BFCT2ACTVDP" runat="server" style="z-index: 1; left: 1019px; position: absolute;top: 0px;" ImageUrl="iMages/FCT2ACTVDP.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BRIGHT" runat="server" style="z-index: 1; left: 1104px; position: absolute;top: 0px;" ImageUrl="iMages/RIGHT.jpg" Height="35px" Width="54px" />
<!-- PAGE-2(上方) -->
        <asp:ImageButton ID="BSTOCK" runat="server" style="z-index: 1; left: 1159px; position: absolute;top: 0px;" ImageUrl="iMages/STOCK.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BALERT" runat="server" style="z-index: 1; left: 1244px; position: absolute;top: 0px;" ImageUrl="iMages/ALERT.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BVDPERR" runat="server" style="z-index: 1; left: 1329px; position: absolute;top: 0px;" ImageUrl="iMages/VDPERROR.jpg" Height="35px" Width="84px" />
        <asp:ImageButton ID="BVDPEDIT" runat="server" style="z-index: 1; left: 1414px; position: absolute;top: 0px;" ImageUrl="iMages/VDPEDIT.jpg" Height="35px" Width="84px" />
<!-- EDI(下方) -->
        <asp:ImageButton ID="BGotoEDI" runat="server" style="z-index: 1; left: 1195px; position: absolute;top: 664px" ImageUrl="iMages/GotoEDI.png" Height="50px" Width="124px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:ImageButton ID="BCustomer" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 142px; position: absolute;top: 77px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" />
        <asp:ImageButton ID="BExcel" runat="server" style="z-index: 1; left: 165px; position: absolute;top: 314px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />
        <asp:ImageButton ID="BConvert" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 556px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 2列  -->
        <asp:ImageButton ID="BFCTPlan" runat="server" style="z-index: 1; left: 552px; position: absolute;top:77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BLSPlan" runat="server" style="z-index: 1; left: 525px; position: absolute;top: 313px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BReport" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 556px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 3列  -->
        <asp:ImageButton ID="BIPLSPlan" runat="server" style="z-index: 1; left: 912px; position: absolute;top:77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BBULSPlan" runat="server" style="z-index: 1; left: 885px; position: absolute;top: 313px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BBUReport" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 556px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 4列  -->
        <asp:ImageButton ID="BLFLSPlan" runat="server" style="z-index: 1; left: 1245px; position: absolute;top: 77px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BEPEDI" runat="server" style="z-index: 1; left: 1245px; position: absolute;top: 313px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BEDIReport" runat="server" style="z-index: 1; left: 1245px; position: absolute;top: 556px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查資料-狀態                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:HyperLink ID="StsCustomer" runat="server" style="z-index: 1; left: 142px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsExcel" runat="server" style="z-index: 1; left: 165px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsConvert" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 606px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 2列  -->
        <asp:HyperLink ID="StsFCTPlan" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsLSPlan" runat="server" style="z-index: 1; left: 525px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsReport" runat="server" style="z-index: 1; left: 552px; position: absolute;top: 606px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 3列  -->
        <asp:HyperLink ID="StsIPLSPlan" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsBULSPlan" runat="server" style="z-index: 1; left: 885px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsBUReport" runat="server" style="z-index: 1; left: 912px; position: absolute;top: 606px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 4列  -->
        <asp:HyperLink ID="StsLFLSPlan" runat="server" style="z-index: 1; left: 1245px; position: absolute;top: 127px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsEPEDI" runat="server" style="z-index: 1; left: 1245px; position: absolute;top: 364px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsEDIReport" runat="server" style="z-index: 1; left: 1245px; position: absolute;top: 606px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <marquee id="ProExcel" runat="server" style="z-index: 1; left: 8px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProConvert" runat="server" style="z-index: 1; left: 10px; position: absolute;top: 680px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 2列  -->
        <marquee id="ProFCTPlan" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 198px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProLSPlan" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProReport" runat="server" style="z-index: 1; left: 370px; position: absolute;top: 680px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 3列  -->
        <marquee id="ProIPLSPlan" runat="server" style="z-index: 1; left: 730px; position: absolute;top: 198px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProBULSPlan" runat="server" style="z-index: 1; left: 730px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProBUReport" runat="server" style="z-index: 1; left: 730px; position: absolute;top: 680px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 4列  -->
        <marquee id="ProLFLSPlan" runat="server" style="z-index: 1; left: 1090px; position: absolute;top: 198px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProEPEDI" runat="server" style="z-index: 1; left: 1090px; position: absolute;top: 440px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProEDIReport" runat="server" style="z-index: 1; left: 1090px; position: absolute;top: 680px; width: 210px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  客戶選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:DropDownList ID="DCustomerBuyer" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 9px; position: absolute; top: 180px" Width="172px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:CheckBox ID="AtReport" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 125px" Font-Size="9pt" Text="Report" Width="100px" />
        <asp:CheckBox ID="AtIPLSPlan" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 145px" Font-Size="9pt" Text="匯入LS Plan" Width="100px" />
        <asp:CheckBox ID="AtBULSPlan" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 165px" Font-Size="9pt" Text="Buyer Proc." Width="100px" />

        <asp:CheckBox ID="AtConvert" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 185px" Font-Size="9pt" Text="Convert" Width="100px" />
        <asp:CheckBox ID="AtFCTPlan" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 205px" Font-Size="9pt" Text="FCT Plan" Width="100px" />
        <asp:CheckBox ID="AtLSPlan" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 225px" Font-Size="9pt" Text="LS Plan" Width="100px" />

        <asp:CheckBox ID="AtLFLSPlan" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 245px" Font-Size="9pt" Text="最終確定" Width="100px" />
        <asp:CheckBox ID="AtEPEDI" runat="server" style="z-index: 174; left: 209px; position: absolute; top: 265px" Font-Size="9pt" Text="EDI Convert" Width="100px" />
        <!-- ITEM / COLOR CONVERT -->
        <asp:CheckBox ID="AtAddition" runat="server" style="z-index: 174; left: 212px; position: absolute; top: 534px" Font-Size="9pt" Text="Addition" Width="100px" />
        <!-- Office2010 -->
        <asp:CheckBox ID="AtOffice2010" runat="server" style="z-index: 174; left: 240px; position: absolute; top: 94px" Font-Size="9pt" Text="Office2010" Width="82px" Checked="True" />
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
        <asp:TextBox ID="DLastVersion" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- ReportData  -->
        <asp:TextBox ID="DReportFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DReportFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DReportFilePath2" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPLANReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFCT2ACTReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFCT2ACTMMReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFCT2ACTVDPReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DMAT2SEAReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DMAT2MONReportFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DLSInputFilePath" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
