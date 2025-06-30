<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DatePicker.aspx.vb" Inherits="DatePicker" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>行事曆</title>
      <style type="text/css">
		BODY { PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 4px; PADDING-TOP: 0px }
		BODY { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		TABLE { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		TR { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
		TD { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	</style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;<asp:Calendar ID="Calendar1" runat="server" BorderColor="Black" ShowGridLines="True">
            <TodayDayStyle BackColor="#FFCC66" ForeColor="White" />
            <SelectorStyle BackColor="#FFCC66" />
            <NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC" />
            <DayHeaderStyle BackColor="#FFCC66" Height="1px" />
            <SelectedDayStyle BackColor="#CCCCFF" Font-Bold="True" />
            <TitleStyle BackColor="#990000" Font-Bold="True" Font-Size="9pt" ForeColor="#FFFFCC" />
            <OtherMonthDayStyle ForeColor="#CC9966" />
        </asp:Calendar>
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
        <br />
    
    </div>
    </form>
</body>
</html>
