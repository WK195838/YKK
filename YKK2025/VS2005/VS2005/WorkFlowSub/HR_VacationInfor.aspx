<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_VacationInfor.aspx.vb" Inherits="HR_VacationInfor" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>個人休假資訊</title>
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Image ID="Image1" runat="server" ImageUrl="images\VacationInfor_001_120621.jpg" Height="178px" Width="677px" />
        <asp:Image ID="Image2" runat="server" ImageUrl="images\vacationinfor_002_120621.jpg" Height="628px" Width="677px" />        

        <asp:HyperLink ID="LVacationList" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 611px; position: absolute; top: 213px" Target="_blank"
            Width="66px">調閱請假</asp:HyperLink>
        <asp:HyperLink ID="LStopJobList" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 123; left: 275px; position: absolute; top: 213px" Target="_blank"
            Width="66px">留職停薪</asp:HyperLink>


        <br />
        <asp:DropDownList ID="DDivision" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 100; left: 120px; position: absolute;
            top: 129px" Width="225px">
        </asp:DropDownList>
        <asp:Image ID="Image3" runat="server" Height="114px" ImageUrl="images\vacationinfor_003_121221.jpg"
            Width="677px" />
        <asp:DropDownList ID="DName" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 106; left: 456px; position: absolute; top: 129px"
            Width="137px" AutoPostBack="True">
        </asp:DropDownList>
        <asp:Button ID="BSimulate" runat="server" Height="25px"
            Style="z-index: 104; left: 268px; position: absolute; top: 159px" Text="Go" Width="78px" />
        <asp:TextBox ID="DBaseDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 103; left: 120px;
            position: absolute; top: 158px" Width="82px">2008/10/10</asp:TextBox>
        <asp:TextBox ID="DInDate" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 103; left: 120px; position: absolute;
            top: 213px" Width="220px">2008/10/10</asp:TextBox>
        <asp:TextBox ID="DNensi" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 103; left: 455px; position: absolute;
            top: 213px" Width="138px">2008/10/10</asp:TextBox>
        <asp:Button ID="BBaseDate" runat="server" Height="25px"
            Style="z-index: 104; left: 208px; position: absolute; top: 158px" Text="....."
            Width="27px" />
        <asp:TextBox ID="DJobTitle" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="True" Style="z-index: 109; left: 456px;
            position: absolute; top: 164px" Width="131px">DJobTitle</asp:TextBox>
        <asp:TextBox ID="DJobCode" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="True" Style="z-index: 116; left: 595px;
            position: absolute; top: 164px" Width="81px">DJobCode</asp:TextBox>
        <asp:TextBox ID="DA_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 280px" Width="51px">0</asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DA_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 280px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DA_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 280px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DA_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 280px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DA_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 280px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DI_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 314px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DI_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 314px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DI_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 314px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DI_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 314px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DI_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 314px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DY_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 348px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DY_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 348px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DY_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 348px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DY_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 348px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DY_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 348px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DS_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 382px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DS_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 382px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DS_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 382px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DS_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 382px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DS_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 382px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DZ_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 434px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DZ_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 434px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DZ_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 434px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DZ_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 434px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DZ_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 434px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DB_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 469px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DB_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 469px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DB_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 469px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DB_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 469px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DB_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 469px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DG_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 503px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DG_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 503px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DG_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 503px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DG_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 503px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DG_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 503px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DM_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 537px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DM_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 537px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DM_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 537px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DM_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 537px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DM_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 537px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DX_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 629px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DX_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 629px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DX_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 629px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DX_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 629px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DX_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 629px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DP_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 663px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DP_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 663px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DP_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 663px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DP_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 663px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DP_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 663px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DQ_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 697px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DQ_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 697px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DQ_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 697px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DQ_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 697px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DQ_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 697px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DH_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 731px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DH_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 731px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DH_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 731px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DH_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 731px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DH_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 731px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DE_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 765px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DE_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px; position: absolute; text-align: right;
            top: 765px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DE_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px; position: absolute; text-align: right;
            top: 765px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DE_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 765px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DE_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 765px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DSum_Lastyear" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 184px;
            position: absolute; text-align: right; top: 799px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DSum_Base" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 247px;
            position: absolute; text-align: right; top: 799px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DSum_Sum" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 310px;
            position: absolute; text-align: right; top: 799px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DSum_Finish" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 373px;
            position: absolute; text-align: right; top: 799px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DSum_Vacation" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" ReadOnly="True" Style="z-index: 126; left: 436px;
            position: absolute; text-align: right; top: 799px" Width="51px">0</asp:TextBox>
        <asp:TextBox ID="DEmpID" runat="server" BackColor="#C0FFC0" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 108; left: 595px; position: absolute;
            top: 131px" Width="81px">DEmpID</asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFC0" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 107; left: 120px; position: absolute;
            top: 98px" Width="220px">DDate</asp:TextBox>
        </div>
    </form>
</body>
</html>
