<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DISPOSALTYPE.aspx.vb" Inherits="DISPOSALTYPE" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>表單</title>
    
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
        &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;<br />
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <br />
        <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 100; left: 30px; position: absolute;
            top: 274px" BorderColor="White" BorderStyle="None"></asp:TextBox>
        <asp:ListBox ID="ListBox1" runat="server" BackColor="Yellow" Height="286px" Rows="3"
            SelectionMode="Multiple" Style="z-index: 101; left: 24px; position: absolute;
            top: 26px" Width="230px">
        </asp:ListBox>
        <asp:Button ID="BGo" runat="server" Style="z-index: 102; left: 264px; position: absolute;
            top: 30px" Text="Go" />
        <asp:TextBox ID="TextBox2" runat="server" Style="z-index: 103; left: 28px; position: absolute;
            top: 312px" BorderColor="White" BorderStyle="None"></asp:TextBox>
        &nbsp;&nbsp;
    
    </div>
    </form>
</body>
</html>
