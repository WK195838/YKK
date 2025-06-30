<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NewCommission.aspx.vb" Inherits="NewCommission" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
	<HEAD>
		<title>新委託</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
		        <!-- ****************************************************************************************** -->
		        <!-- ** Main Images
		        <!-- ****************************************************************************************** -->
				<asp:image id="DSheet1" style="Z-INDEX: 100; LEFT: 16px; POSITION: absolute; TOP: 32px" runat="server"
					ImageUrl="Images\ITMMain_250105.png" Width="1240px" Height="467px"></asp:image>
		        <!-- ****************************************************************************************** -->
		        <!-- ** Download
		        <!-- ****************************************************************************************** -->
                <asp:HyperLink ID="HLink_04" runat="server" Font-Bold="True" Font-Size="Small"
                    ForeColor="Blue" Height="24px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/List/AllItems.aspx?env=WebViewList"
                    Style="z-index: 104; left: 600px; position: absolute; top: 120px; text-align: center"
                    Target="_blank" Width="152px">數據變更申請表(WORD)</asp:HyperLink>

                <asp:HyperLink ID="HLink_05" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 600px; position: absolute; text-align:center;
                    top: 148px" Target="_blank" Width="152px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/List/AllItems.aspx?env=WebViewList" 
                    Font-Bold="True" Font-Size="Small">VB/RCS申請表(WORD)</asp:HyperLink>

                <asp:HyperLink ID="HLink_06" runat="server" Font-Bold="True" Font-Size="Small" ForeColor="Blue"
                    Height="24px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/List/AllItems.aspx?env=WebViewList"
                    Style="z-index: 104; left: 600px; position: absolute; top: 176px; text-align: center"
                    Target="_blank" Width="152px">系統變更申請表(WORD)</asp:HyperLink>
                    
                <asp:HyperLink ID="HLink_21" runat="server" Font-Bold="True" Font-Size="Small"
                    ForeColor="Blue" Height="16px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/List/AllItems.aspx?env=WebViewList"
                    Style="z-index: 104; left: 600px; position: absolute; top: 208px; text-align: left"
                    Target="_blank" Width="152px">新特別要求追加表</asp:HyperLink>

            <asp:TextBox ID="DMemo1" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Red"
                Height="28px" Style="z-index: 126; left: 408px; position: absolute;
                top: 48px; text-align: left" Width="832px" MaxLength="50" Font-Bold="False" Font-Size="Medium" Font-Underline="True">［注意］：特別要求追加時需填［新特別要求追加表］並附上申請單上，無此表時將無法處理 (3/10 13:30開始實施)</asp:TextBox>

		        <!-- ****************************************************************************************** -->
		        <!-- ** Manual
		        <!-- ****************************************************************************************** -->
                <asp:HyperLink ID="HLink_ITM" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 824px; position: absolute; text-align:left;
                    top: 120px" Target="_blank" Width="160px" NavigateUrl="https://ykkglobal.sharepoint.com/sites/asia_twn_discuss_cpu/Lists/List/DispForm.aspx?ID=189&e=Sc7LIp"  
                    Font-Bold="True" Font-Size="Small">使用手冊(New 250105)</asp:HyperLink>

                <asp:HyperLink ID="HLink_Purpose" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 824px; position: absolute; text-align:left;
                    top: 148px" Target="_blank" Width="56px" NavigateUrl="http://10.245.1.10/ITM/Images/Purposedescription.png" 
                    Font-Bold="True" Font-Size="Small">用途說明</asp:HyperLink>

                <asp:HyperLink ID="HLink_DevelopNotice" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 824px; position: absolute; text-align:left;
                    top: 176px" Target="_blank" Width="208px" NavigateUrl="http://10.245.1.10/ITM/Images/Depoment_Notice.png" 
                    Font-Bold="True" Font-Size="Small">RCS/Excel/Sys.Change需要文件</asp:HyperLink>
                    
		        <!-- ****************************************************************************************** -->
		        <!-- ** Q&A
		        <!-- ****************************************************************************************** -->
                <asp:HyperLink ID="HLQAForm" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 1040px; position: absolute; text-align:left;
                    top: 120px" Target="_blank" Width="96px" NavigateUrl="https://forms.office.com/r/uUpwSwLY0W" 
                    Font-Bold="True" Font-Size="Small">QA Form</asp:HyperLink>

                <asp:HyperLink ID="HLQATeams" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 1040px; position: absolute; text-align:left;
                    top: 148px" Target="_blank" Width="56px" NavigateUrl="https://teams.microsoft.com/dl/launcher/launcher.html?url=%2F_%23%2Fl%2Fchannel%2F19%3Aee77d6ac9aa349e3ad1534ee40e379ee%40thread.tacv2%2F090-ITM%3FgroupId%3D22ac1edd-87b7-4276-b27c-4fe24bf5d391%26tenantId%3D51a2e163-d3e1-40a6-b825-26a4b0b38310&type=channel&deeplinkId=60db7e61-cce3-45fb-abd3-4a739d7335cd&directDl=true&msLaunch=true&enableMobilePage=true&suppressPrompt=true" 
                    Font-Bold="True" Font-Size="Small">Teams</asp:HyperLink>

                <asp:HyperLink ID="HLQATeamsUser" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 1040px; position: absolute; text-align:left;
                    top: 176px" Target="_blank" Width="112px" NavigateUrl="https://forms.office.com/r/uUpwSwLY0W" 
                    Font-Bold="True" Font-Size="Small">Teams權限追加</asp:HyperLink>

                <asp:HyperLink ID="HLQAITM" runat="server" 
                    ForeColor="Blue" Height="24px" Style="z-index: 104; left: 1040px; position: absolute; text-align:left;
                    top: 208px" Target="_blank" Width="112px" NavigateUrl="https://ykkglobal.sharepoint.com/sites/asia_twn_discuss_cpu/Lists/List/DispForm.aspx?ID=193&e=4XU0XS" 
                    Font-Bold="True" Font-Size="Small">ITM思考背景</asp:HyperLink>

		        <!-- ****************************************************************************************** -->
		        <!-- ** Connect
		        <!-- ****************************************************************************************** -->
                <asp:HyperLink ID="HLink_01" runat="server" Font-Bold="True" Font-Size="Larger"
                    ForeColor="Blue" Height="24px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/List/AllItems.aspx?env=WebViewList"
                    Style="z-index: 104; left: 208px; position: absolute; top: 368px; text-align: center"
                    Target="_blank" Width="96px">Connect..</asp:HyperLink>
                <asp:HyperLink ID="HLink_02" runat="server" Font-Bold="True" Font-Size="Larger" ForeColor="Blue"
                    Height="24px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/VBRCS/AllItems.aspx?env=WebViewList"
                    Style="z-index: 104; left: 488px; position: absolute; top: 368px; text-align: center"
                    Target="_blank" Width="88px">Connect..</asp:HyperLink>
                <asp:HyperLink ID="HLink_03" runat="server" Font-Bold="True" Font-Size="Larger" ForeColor="Blue"
                    Height="16px" NavigateUrl="https://ykkglobal-my.sharepoint.com/personal/powerauto_twn_ykk_com/Lists/List1/AllItems.aspx?env=WebViewList"
                    Style="z-index: 104; left: 840px; position: absolute; top: 368px; text-align: center"
                    Target="_blank" Width="88px">Connect..</asp:HyperLink>

                <asp:HyperLink ID="HLink_11" runat="server" Font-Bold="True" Font-Size="Larger" ForeColor="Blue"
                    Height="16px" NavigateUrl="https://ykkglobal.sharepoint.com/:x:/r/teams/ISOS/_layouts/15/Doc.aspx?sourcedoc=%7B5BC51FC7-2B5A-4940-AF48-CFE34C2F8834%7D&file=24%25u5e74%25u5ea6%25u53f0%25u6e7eIT%25u30bf%25u30b9%25u30af%25u30ea%25u30b9%25u30c8_2025.xlsx&wdLOR=c48F30B16-C501-47FC-825D-E010F14FE42D&action=default&mobileredirect=true"
                    Style="z-index: 104; left: 984px; position: absolute; top: 340px; text-align: center"
                    Target="_blank" Width="88px">Connect..</asp:HyperLink>

                <asp:HyperLink ID="HLink_12" runat="server" Font-Bold="True" Font-Size="Larger" ForeColor="Blue"
                    Height="16px" NavigateUrl="https://ykkglobal.sharepoint.com/:x:/r/teams/ISOS/_layouts/15/Doc.aspx?sourcedoc=%7BFFEB9709-7465-4858-BFC2-FADDD326B687%7D&file=IT-SD%20%25u8ab2%25u984c%25u4e00%25u89bd%25u8868.xlsx&wdLOR=c880C9177-FC72-48CE-8D3C-A7FE9C69B0B1&action=default&mobileredirect=true"
                    Style="z-index: 104; left: 1128px; position: absolute; top: 340px; text-align: center"
                    Target="_blank" Width="88px">Connect..</asp:HyperLink>
            </FONT>
         </form>
	</body>
</HTML>
