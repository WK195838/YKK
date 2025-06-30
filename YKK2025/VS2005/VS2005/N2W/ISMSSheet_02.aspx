<%@ Page Language="VB" Debug="true" CodeFile="ISMSSheet_02.aspx.vb" Inherits="ISMSSheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ISMS 資訊工作日誌申請單</title>
	  
		<script language="javascript" type="text/javascript">
		
		function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}

	function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			

			
            function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    btn.disabled="disabled";
				    var answer = confirm("是否執行作業嗎？ 請確認....");
				    if (answer) {
                        document.forms[0].__EVENTTARGET.value = btn.name;
                        document.forms[0].__EVENTARGUMENT.value = '';
                        document.forms[0].submit();
				    }                    
                    else    {
                        btn.disabled="";
                        return false;   
                    }				    
                }
                else    {
                    return false;
                }
            }
		</script>
	
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/ISMSSheet_1.jpg" style="z-index: 1; left: 7px; position: absolute;top: 7px" />
        &nbsp; &nbsp;
        <input id="DCheckDate" runat="server" name="DCheckDate" style="z-index: 132; left: 136px;
            width: 81px; color: black; border-top-style: groove; border-right-style: groove;
            border-left-style: groove; position: absolute; top: 176px; background-color: #c0ffff;
            border-bottom-style: groove" type="text" />
        &nbsp; &nbsp;
        <asp:DropDownList ID="DEroom" runat="server" BackColor="#C0FFFF" Height="20px"
            Style="z-index: 134; left: 240px; position: absolute; top: 248px"
            Width="456px">
        </asp:DropDownList><asp:DropDownList ID="DPlace" runat="server" BackColor="#C0FFFF" Height="20px"
            Style="z-index: 134; left: 240px; position: absolute; top: 214px"
            Width="456px">
        </asp:DropDownList>
        <asp:DropDownList ID="DITName" runat="server" BackColor="#C0FFFF" Height="20px"
            Style="z-index: 134; left: 472px; position: absolute; top: 176px"
            Width="224px">
        </asp:DropDownList>
        <asp:TextBox ID="DState" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="56px" MaxLength="7" Style="z-index: 126;
            left: 136px; position: absolute; top: 376px; text-align: left" TextMode="MultiLine"
            Width="560px"></asp:TextBox>
        <asp:TextBox ID="DTemp" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 244px; position: absolute; top: 280px; text-align: left" Width="112px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DHumidity" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 244px; position: absolute; top: 312px; text-align: left" Width="112px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DOther" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 244px; position: absolute; top: 344px; text-align: left" Width="448px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DCause" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="56px" MaxLength="7" Style="z-index: 126;
            left: 136px; position: absolute; top: 440px; text-align: left" Width="560px" TextMode="MultiLine"></asp:TextBox>
        <asp:TextBox ID="DDeal" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="56px" MaxLength="7" Style="z-index: 126;
            left: 136px; position: absolute; top: 506px; text-align: left" TextMode="MultiLine"
            Width="560px"></asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="56px" MaxLength="7" Style="z-index: 126;
            left: 136px; position: absolute; top: 576px; text-align: left" TextMode="MultiLine"
            Width="560px"></asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 136px; position: absolute; top: 112px; text-align: left" Width="220px"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 472px;
            position: absolute; top: 112px">DNo</asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 136px;
            position: absolute; top: 144px; text-align: left" Width="220px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DAppName" runat="server" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 472px; position: absolute; top: 144px; text-align: left" Width="220px"></asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 132; left: 768px; position: absolute;
               top: 300px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 133; left: 769px; position: absolute; top: 337px"
               Width="190px"></asp:TextBox>

        <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
        &nbsp; &nbsp; &nbsp;
    </div>
    </form>
</body>
</html>
