<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPD_YKKGroupCopySheet_02.aspx.vb" Inherits="SPD_YKKGroupCopySheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>姊妹社圖面複製履歷</title>
    
    <script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			   function GetMapNo()
{
        
    window.open('MapNoList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
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
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp;&nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/YKKGroupCopySheet_01.jpg"
            Style="z-index: 110; left: 4px; position: absolute; top: 6px" />
        <asp:TextBox ID="DYKKGroup" runat="server" BackColor="#C0FFFF" Style="z-index: 123;
            left: 213px; position: absolute; top: 265px" Width="574px"></asp:TextBox>
        <asp:TextBox ID="DMapNo" runat="server" BackColor="#C0FFFF" Style="z-index: 111;
            left: 211px; position: absolute; top: 171px" Width="188px"></asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DProvideDate" runat="server" BackColor="#C0FFFF" Style="z-index: 115;
            left: 604px; position: absolute; top: 393px" Width="184px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="52px" MaxLength="100" Style="z-index: 113; left: 212px;
            position: absolute; top: 426px; text-align: left" TextMode="MultiLine" Width="580px"></asp:TextBox>
        <asp:TextBox ID="DSliderCode" runat="server" BackColor="#C0FFFF" Style="z-index: 114;
            left: 211px; position: absolute; top: 201px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DWaveCode" runat="server" BackColor="#C0FFFF" Style="z-index: 115;
            left: 601px; position: absolute; top: 202px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" Style="z-index: 116;
            left: 601px; position: absolute; top: 105px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DPerson" runat="server" BackColor="#C0FFFF" Style="z-index: 117;
            left: 601px; position: absolute; top: 136px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="#C0FFFF" Style="z-index: 118;
            left: 601px; position: absolute; top: 168px" Width="190px"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="DNo" runat="server" BackColor="#C0FFFF" Style="z-index: 119;
            left: 213px; position: absolute; top: 106px" Width="190px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="#C0FFFF" Style="z-index: 120;
            left: 211px; position: absolute; top: 136px" Width="190px"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DForcast" runat="server" BackColor="#C0FFFF" Style="z-index: 122;
            left: 210px; position: absolute; top: 395px" Width="194px"></asp:TextBox>
        <asp:TextBox ID="DCopyReason" runat="server" BackColor="#C0FFFF" Style="z-index: 123;
            left: 419px; position: absolute; top: 361px" Width="367px"></asp:TextBox>
        &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:CheckBox ID="DCopyCheck1" runat="server" Style="z-index: 124; left: 214px; position: absolute;
            top: 300px" AutoPostBack="True" />
        <asp:CheckBox ID="DCopyCheck2" runat="server" Style="z-index: 125; left: 214px; position: absolute;
            top: 333px" />
        <asp:CheckBox ID="DCopyCheck3" runat="server" Style="z-index: 126; left: 214px; position: absolute;
            top: 365px" />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:HyperLink ID="LFormsno" runat="server" Style="z-index: 127; left: 211px; position: absolute;
            top: 235px" Width="103px" Target="_blank">原委託</asp:HyperLink>
        &nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
            <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 129; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
        &nbsp;
            
        </div>
    </form>
   

</body>
</html>


