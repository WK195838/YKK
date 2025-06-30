<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NewReport.aspx.vb" Inherits="NewReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新報告</title>
	<script language="javascript" type="text/javascript">
			function CalendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'CalendarPopup','width=250,height=190,resizable=yes');
			}
			
	</script>
</head>
<body>
    <form id="Form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Image ID="DSheet1" runat="server" Height="500px" ImageUrl="Images\NewReport.png"
            Style="z-index: 104; left: 8px; position: absolute; top: 3px" Width="645px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  表單欄位                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DReportDate" runat="server" BackColor="GreenYellow" Style="z-index: 108; left: 147px; position: absolute;
            top: 94px" Width="143px" BorderStyle="Groove" ForeColor="Blue" Height="20px" MaxLength="10">DReportDate</asp:TextBox>

        <asp:TextBox ID="DID" runat="server" style="Z-INDEX: 107; LEFT: 462px; POSITION: absolute; TOP: 94px" Width="170px" BackColor="LightGray" BorderStyle="Groove" 
            ForeColor="Blue" Height="20px" ReadOnly="True" MaxLength="10" >DID</asp:TextBox>

        <asp:TextBox ID="DName" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 107; left: 147px;
            position: absolute; top: 133px" Width="176px" MaxLength="20">DName</asp:TextBox>


        <asp:TextBox ID="DDivision" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 107; left: 462px;
            position: absolute; top: 133px" Width="176px" MaxLength="20">DDivision</asp:TextBox>
                    
        <asp:DropDownList ID="DSection" runat="server" BackColor="GreenYellow" ForeColor="Blue"
            Style="z-index: 112; left: 147px; position: absolute; top: 176px" Width="304px">
            <asp:ListItem Selected="True"></asp:ListItem>
            <asp:ListItem>INFRA</asp:ListItem>
            <asp:ListItem>WINGS</asp:ListItem>
            <asp:ListItem>WEB</asp:ListItem>
            <asp:ListItem>全般</asp:ListItem>
        </asp:DropDownList>
        
        <asp:DropDownList ID="DType" runat="server" BackColor="GreenYellow" ForeColor="Blue"
            Style="z-index: 112; left: 147px; position: absolute; top: 216px" Width="304px">
            <asp:ListItem Selected="True"></asp:ListItem>
            <asp:ListItem>10.Q&A</asp:ListItem>
            <asp:ListItem>20.運用</asp:ListItem>
            <asp:ListItem>30.設計</asp:ListItem>
            <asp:ListItem>31.開發</asp:ListItem>
            <asp:ListItem>32.測試</asp:ListItem>
            <asp:ListItem>40.支援</asp:ListItem>
            <asp:ListItem>90.其他</asp:ListItem>
            <asp:ListItem>99.休假</asp:ListItem>
        </asp:DropDownList>

        <asp:TextBox ID="DContent" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="96px" Style="z-index: 107; left: 147px;
            position: absolute; top: 256px" Width="488px"  TextMode="MultiLine" >DContent</asp:TextBox>

        <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="96px" Style="z-index: 108; left: 147px; position: absolute;
            top: 376px" Width="484px"  TextMode="MultiLine">DRemark</asp:TextBox>

        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  按鈕                                                                                ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->&nbsp;
        <asp:Button ID="BReportDate" runat="server" Height="25px" Style="z-index: 104; left: 296px;
            position: absolute; top: 94px" Text="....." Width="25px" />

        <asp:Button ID="BSave" runat="server" Text="儲存" style="Z-INDEX: 114; LEFT: 544px; POSITION: absolute; TOP: 512px" Width="100px" />
        <asp:Button ID="BDel" runat="server" Text="刪除" style="Z-INDEX: 114; LEFT: 440px; POSITION: absolute; TOP: 512px" Width="100px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查欄位                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:RequiredFieldValidator ID="DReportDateRqd" runat="server" ErrorMessage="需輸入日期" Style="z-index: 108; left: 16px; position: absolute;
            top: 512px" Width="410px" ControlToValidate="DReportDate"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="DNameRqd" runat="server" ErrorMessage="需輸入姓名" Style="z-index: 108; left: 16px; position: absolute;
            top: 512px" Width="410px" ControlToValidate="DName"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="DDivisionRqd" runat="server" ErrorMessage="需輸入GROUP" Style="z-index: 108; left: 16px; position: absolute;
            top: 512px" Width="410px" ControlToValidate="DDivision"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="DSectionRqd" runat="server" ErrorMessage="需輸入TEAM" Style="z-index: 108; left: 16px; position: absolute;
            top: 512px" Width="410px" ControlToValidate="DSection"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="DTypeRqd" runat="server" ErrorMessage="需輸入作業" Style="z-index: 108; left: 16px; position: absolute;
            top: 512px" Width="410px" ControlToValidate="DType"></asp:RequiredFieldValidator>
        <asp:RequiredFieldValidator ID="DContentRqd" runat="server" ErrorMessage="需輸入內容" Style="z-index: 108; left: 16px; position: absolute;
            top: 512px" Width="410px" ControlToValidate="DContent"></asp:RequiredFieldValidator>
    </div>
    </form>
</body>


</html>
