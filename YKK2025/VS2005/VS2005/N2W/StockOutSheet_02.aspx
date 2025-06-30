<%@ Page Language="vb" AutoEventWireup="false" Inherits="StockOutSheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="StockOutSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>發送出庫申請單</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			               function GetDep()
{
        
    window.open('DepList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	
	
	   function GetSPEC()
{
        
    window.open('StockList.aspx?','','status=0,toolbar=0,width=800,height=650,top=10,resizable=yes');
   
}	

   function GetSPEC1(strField,strField1)
{
        
    window.open('StockList.aspx?pStep=10&field=' + strField+'&field1='+strField1,'','status=0,toolbar=0,width=800,height=650,top=10,resizable=yes');
   
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

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/StockOut_01.jpg" style="z-index: 1; left: 16px; position: absolute;
                   top: 16px" />
               <asp:TextBox ID="DType" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 232px; position: absolute;
                   top: 200px; text-align: left" Width="140px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;
               <asp:TextBox ID="DItemCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 232px; position: absolute;
                   top: 232px; text-align: left" Width="56px"></asp:TextBox>
               <asp:TextBox ID="DColor" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 488px; position: absolute;
                   top: 232px; text-align: left" Width="120px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" Style="z-index: 126; left: 232px; position: absolute;
                   top: 128px; text-align: left" Width="140px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 488px; position: absolute; top: 136px" Width="120px">DNo</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 480px;
                   position: absolute; top: 168px; text-align: left" Width="136px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           </div>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DQty" runat="server"  onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 232px;
               position: absolute; top: 264px; text-align: left" Width="56px"></asp:TextBox>
           <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 232px; position: absolute;
               top: 168px; text-align: left" Width="140px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DTNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 132; left: 952px; position: absolute;
               top: 96px" Width="190px"></asp:TextBox>
           <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="328px" ScrollBars="Auto"
               Style="left: 24px; position: absolute; top: 320px" Width="632px">
               &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
               BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px"
               CellPadding="3" DataKeyNames="stockno" Font-Size="9pt" ForeColor="Black" GridLines="Vertical"
               PageSize="100" Style="z-index: 103; left: 8px; position: absolute; top: 16px" Height="16px" Width="600px">
               <Columns>
                   <asp:BoundField DataField="stockno" HeaderText="棧板號">
                       <HeaderStyle Height="20px" />
                   </asp:BoundField>
               </Columns>
               <FooterStyle BackColor="#CCCCCC" />
               <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
               <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
               <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                   VerticalAlign="Middle" />
               <AlternatingRowStyle BackColor="#CCCCCC" />
           </asp:GridView>
           </asp:Panel>
           <asp:TextBox ID="DCODE" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 952px; position: absolute;
               top: 272px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="DPaletNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 952px; position: absolute;
               top: 232px" Width="190px"></asp:TextBox>
  	 
      </div>
    </form>
</body>
</html>
