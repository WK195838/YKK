<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_WorkRatioInfor.aspx.vb" Inherits="HR_WorkRatioInfor" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>個人出勤率</title>
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
function IMG1_onclick() {

}

		</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <br />
        <asp:DropDownList ID="DDivision" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 100; left: 466px; position: absolute;
            top: 45px" Width="197px">
        </asp:DropDownList>
        <asp:DropDownList ID="DName" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 101; left: 124px; position: absolute; top: 77px"
            Width="137px" AutoPostBack="True">
        </asp:DropDownList>
        <asp:TextBox ID="DJobTitle" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="True" Style="z-index: 102; left: 466px;
            position: absolute; top: 80px" Width="93px">DJobTitle</asp:TextBox>
        <asp:TextBox ID="DYear" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="True" Style="z-index: 103; left: 466px;
            position: absolute; top: 11px" Width="93px"></asp:TextBox>
        <asp:TextBox ID="DBaseDate" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="True" Style="z-index: 104; left: 124px;
            position: absolute; top: 44px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DJobCode" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="True" Style="z-index: 105; left: 577px;
            position: absolute; top: 80px" Width="81px">DJobCode</asp:TextBox>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/HR_WorkRatioInfor_01.PNG"
            Style="z-index: 99; left: 2px; position: absolute; top: -1px" />
        <asp:TextBox ID="DEmpID" runat="server" BackColor="#C0FFC0" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 107; left: 274px; position: absolute;
            top: 77px" Width="61px">DEmpID</asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFC0" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 108; left: 124px; position: absolute;
            top: 11px" Width="213px">DDate</asp:TextBox>
        <asp:HyperLink ID="LOtherDescription" runat="server" Font-Size="12pt" Height="8px"
            NavigateUrl="BoardEdit.aspx" Style="z-index: 109; left: 618px; position: absolute;
            top: 483px" Target="_blank" Visible="False" Width="66px">其它說明</asp:HyperLink>
        <asp:HyperLink ID="LVacationList" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 110; left: 617px; position: absolute; top: 459px" Target="_blank"
            Visible="False" Width="66px">休假調閱</asp:HyperLink>
        <asp:TextBox ID="DYVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 110; left: 129px;
            position: absolute; top: 208px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DSVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 111; left: 129px;
            position: absolute; top: 242px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DZVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 112; left: 129px;
            position: absolute; top: 273px" Width="46px"></asp:TextBox>
        <asp:Label ID="DSum" runat="server" ForeColor="Blue" Style="z-index: 113; left: 129px;
            position: absolute; top: 503px" Width="335px"></asp:Label>
        <asp:TextBox ID="DMVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 114; left: 129px;
            position: absolute; top: 307px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DXVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 115; left: 129px;
            position: absolute; top: 341px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DPVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 116; left: 129px;
            position: absolute; top: 375px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DQVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 117; left: 129px;
            position: absolute; top: 406px" Width="46px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DIVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 118; left: 231px;
            position: absolute; top: 178px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DYVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 119; left: 231px;
            position: absolute; top: 208px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DSVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 120; left: 231px;
            position: absolute; top: 242px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DZVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 121; left: 231px;
            position: absolute; top: 273px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DMVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 122; left: 231px;
            position: absolute; top: 307px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DXVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 123; left: 231px;
            position: absolute; top: 341px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DPVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 124; left: 231px;
            position: absolute; top: 375px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DQVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 125; left: 231px;
            position: absolute; top: 406px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DZ1M" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 126; left: 231px;
            position: absolute; top: 443px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DX2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 127; left: 231px;
            position: absolute; top: 475px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DYearDay" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 128; left: 465px;
            position: absolute; top: 178px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DVacationday" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 129; left: 465px;
            position: absolute; top: 208px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DWorkRate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 130; left: 592px;
            position: absolute; top: 320px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DY2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 131; left: 465px;
            position: absolute; top: 242px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DY3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 132; left: 465px;
            position: absolute; top: 273px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DAWorkday" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 133; left: 231px;
            position: absolute; top: 146px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DBworkday" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 134; left: 465px;
            position: absolute; top: 146px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DIVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 135; left: 129px;
            position: absolute; top: 178px" Width="46px"></asp:TextBox>
        </div>
    </form>
</body>
</html>
