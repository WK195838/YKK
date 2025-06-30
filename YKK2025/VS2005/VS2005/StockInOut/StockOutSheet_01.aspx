<%@ Page Language="vb" AutoEventWireup="false" Inherits="StockOutSheet_01" aspCompat="True" EnableEventValidation = "false"  CodeFile="StockOutSheet_01.aspx.vb" %>
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

   function GetSPEC1(strField,strField1,strField2,strField3)
{
        
    window.open('StockList.aspx?pStep=10&field=' + strField+'&field1='+strField1+'&field2='+strField2+'&field3='+strField3,'','status=0,toolbar=0,width=850,height=650,top=10,resizable=yes');
   
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
            



 function UploadFile(fileUpload) {
        if (fileUpload.value != '') {
            document.getElementById("<%=btn_upload.ClientID %>").click();
        }
    }


		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/StockOut_04.jpg" style="z-index: 1; left: 8px; position: absolute;
                   top: 0px" />
               <asp:TextBox ID="DQty1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" onkeyup="if(isNaN(value))execCommand('undo')"
                   Style="z-index: 126; left: 456px; position: absolute; top: 280px; text-align: left"
                   Visible="False" Width="56px"></asp:TextBox>
               <asp:FileUpload ID="DAttachfile" runat="server" Height="24px" Style="z-index: 100;
                   left: 224px; position: absolute; top: 240px; background-color: #c0ffff" Visible="False"
                   Width="400px" />
               <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 224px; position: absolute;
                   top: 312px; text-align: left" Width="376px"></asp:TextBox>
               &nbsp;
               <asp:DropDownList ID="DType" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 102;
                   left: 224px; position: absolute; top: 184px" Visible="False" Width="144px" AutoPostBack="True">
               </asp:DropDownList>
               <input id="BSPEC" runat="server" style="z-index: 152; left: 320px; width: 24px; position: absolute;
                   top: 216px" type="button" value="..." visible="false" />
               &nbsp;&nbsp;&nbsp;&nbsp;
               <asp:TextBox ID="DItemCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 224px; position: absolute;
                   top: 216px; text-align: left" Width="88px" Visible="False"></asp:TextBox>
               <asp:TextBox ID="DColor" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 456px; position: absolute;
                   top: 216px; text-align: left" Width="120px" Visible="False"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" Style="z-index: 126; left: 224px; position: absolute;
                   top: 112px; text-align: left" Width="112px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 464px; position: absolute; top: 120px" Width="120px">DNo</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 456px;
                   position: absolute; top: 152px; text-align: left" Width="136px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 12px;
                   position: absolute; top: 1034px" visible="false" />
               <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 132; left: 59px; position: absolute; top: 965px"
                   TextMode="MultiLine" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
                   Style="z-index: 133; left: 243px; position: absolute; top: 1039px" Visible="False"
                   Width="352px"></asp:TextBox>
               <asp:DropDownList ID="DReasonCode" runat="server" BackColor="Yellow" Height="20px"
                   Style="z-index: 134; left: 167px; position: absolute; top: 1042px" Visible="False"
                   Width="64px">
                   <asp:ListItem>01</asp:ListItem>
                   <asp:ListItem>02</asp:ListItem>
               </asp:DropDownList>
               <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 135; left: 171px; position: absolute; top: 1072px"
                   Visible="False" Width="424px"></asp:TextBox>
               <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
                   Style="z-index: 137; left: 14px; position: absolute; top: 1157px" Text="核定履歷"></asp:Label>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                   Style="z-index: 136; left: 13px; position: absolute; top: 1179px" Width="780px">
                   <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
                   <Columns>
                       <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                           <ItemStyle HorizontalAlign="Left" />
                       </asp:BoundField>
                       <asp:BoundField DataField="DecideName" HeaderText="擔當" />
                       <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" />
                       <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" />
                       <asp:BoundField DataField="StsDesc" HeaderText="處理結果" />
                       <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                           <ItemStyle HorizontalAlign="Left" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Description" HeaderText="核定時間">
                           <ItemStyle HorizontalAlign="Left" />
                       </asp:BoundField>
                   </Columns>
                   <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                       VerticalAlign="Middle" />
               </asp:GridView>
               <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
                   left: 13px; position: absolute; top: 959px" />
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 114; left: 354px; position: absolute; top: 1144px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 115; left: 445px; position: absolute; top: 1144px" Text="NG2"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 116; left: 537px; position: absolute; top: 1144px" Text="NG1"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 117; left: 629px; position: absolute; top: 1144px" Text="OK"
                   UseSubmitBehavior="false" Width="80px" />
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
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 224px;
               position: absolute; top: 280px; text-align: left" Width="56px" Visible="False"></asp:TextBox>
           <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 224px; position: absolute;
               top: 152px; text-align: left" Width="112px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DTNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 132; left: 952px; position: absolute;
               top: 96px" Width="190px"></asp:TextBox>
           <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="256px" ScrollBars="Auto"
               Style="left: 32px; position: absolute; top: 360px" Width="632px">
               &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
               BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px"
               CellPadding="3" DataKeyNames="stockno" Font-Size="9pt" ForeColor="Black" GridLines="Vertical"
               PageSize="100" Style="z-index: 103; left: 8px; position: absolute; top: 8px" Height="16px" Width="600px">
               <Columns>
                   <asp:BoundField DataField="stockno" HeaderText="棧板號">
                       <HeaderStyle Height="20px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="Qty" HeaderText="Qty" />
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
           <asp:GridView ID="GridView3" runat="server" Style="z-index: 120; left: 912px; position: absolute;
               top: 360px" Visible="False">
           </asp:GridView>
           <asp:RadioButtonList ID="rbHDR" runat="server" Style="z-index: 121; left: 1008px;
               position: absolute; top: 512px" Visible="False">
               <asp:ListItem Selected="True" Text="Yes" Value="Yes"></asp:ListItem>
               <asp:ListItem Text="No" Value="No"></asp:ListItem>
           </asp:RadioButtonList>
           <asp:Button ID="btn_upload" runat="server" Height="32px" OnClick="btn_upload_Click"
               Style="display: none; z-index: 123; left: 936px; position: absolute; top: 288px" />
           <asp:TextBox ID="DBoxQty" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" onkeyup="if(isNaN(value))execCommand('undo')" Style="z-index: 126;
               left: 544px; position: absolute; top: 280px; text-align: left" Visible="False"
               Width="32px"></asp:TextBox>
  	 
      </div>
    </form>
</body>
</html>
