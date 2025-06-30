<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_NewColorinqPriority" aspCompat="True"  EnableEventValidation = "false"  CodeFile="DTMW_NewColorinqPriority.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新色依賴書優先度調閱資料</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
   function GetStsPicker()
{
        
    window.open('StsPicker.aspx?','','status=0,toolbar=0,width=250,height=190,top=10,resizable=yes');
   
}	 

			
   function GetFormPicker()
{
        
    window.open('FormPicker.aspx?','','status=0,toolbar=0,width=280,height=220,top=10,resizable=yes');
   
}

   function GetDivisionPicker()
{
        
    window.open('DivisionPicker.aspx?','','status=0,toolbar=0,width=250,height=190,top=10,resizable=yes');
   
}


		    function Button(F, MSG) {
				//alert(F);
				document.cookie="RunBOK=False";
				document.cookie="RunBNG1=False";
				document.cookie="RunBNG2=False";
				document.cookie="RunBSAVE=False";

				answer = confirm("是否執行[" + MSG + "]作業嗎？ 請確認....");
				if (answer) {
					//OK Button
					//FOK=document.getElementById("BOK");
					//if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					//FNG1=document.getElementById("BNG1");
					//if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					//FNG2=document.getElementById("BNG2");
					//if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					//FSAVE=document.getElementById("BSAVE");
					//if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
					if (F=="OK")   document.cookie="RunBOK=True";
					if (F=="NG1")  document.cookie="RunBNG1=True";
					if (F=="NG2")  document.cookie="RunBNG2=True";
					if (F=="SAVE") document.cookie="RunBSAVE=True";
				}
			}
		   
		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 448px; POSITION: absolute; TOP: 104px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px" Visible="False"></asp:imagebutton>
           <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False"
               BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px"
               CellPadding="4" Font-Size="9pt" Height="100px" PageSize="15" Style="z-index: 101;
               left: 8px; position: absolute; top: 136px" Width="1448px">
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <SelectedItemStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
               <ItemStyle BackColor="White" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               <Columns>
                   <asp:HyperLinkColumn DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
                       DataTextField="Field1" Target="_blank">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:HyperLinkColumn>
                   <asp:BoundColumn DataField="Field2" HeaderText="狀態" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field3" HeaderText="依賴日" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="CompleteDate" HeaderText="完成日"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field4" HeaderText="依賴表單" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field5" HeaderText="依賴部門" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field6" HeaderText="依賴者" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field7" HeaderText="客戶" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field8" HeaderText="Buyer" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="CustomerColor" HeaderText="客戶色名"></asp:BoundColumn>
                   <asp:BoundColumn DataField="CustomerColorCode" HeaderText="客戶色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="OverSeaYkkcode" HeaderText="海外YKK色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="PantoneCode" HeaderText="PANTONECode"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field9" HeaderText="YKK色別"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field10" HeaderText="YKK色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field11" HeaderText="兼用色拉頭"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field12" HeaderText="兼用色VF上/下止"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field13" HeaderText="兼用色VF上/下止"></asp:BoundColumn>
                   <asp:BoundColumn DataField="NewOldColor" HeaderText="新舊色"></asp:BoundColumn>
                   <asp:BoundColumn HeaderText="對應YKK色號" DataField="ReColorCode"></asp:BoundColumn>
                   <asp:HyperLinkColumn DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}"
                       DataTextField="WorkFlow" HeaderText="履歷" Target="_blank">
                       <HeaderStyle Width="30px" />
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:HyperLinkColumn>
                   <asp:TemplateColumn HeaderText="優先度">
                       <EditItemTemplate>
                           <asp:TextBox runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Priority") %>'></asp:TextBox>
                       </EditItemTemplate>
                       <ItemTemplate>
                           <asp:Label runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Priority") %>'></asp:Label>
                       </ItemTemplate>
                       <HeaderStyle BackColor="Yellow" ForeColor="Red" />
                   </asp:TemplateColumn>
               </Columns>
               <PagerStyle BackColor="#FFFFCC" ForeColor="#993300" HorizontalAlign="Center" Mode="NumericPages"
                   NextPageText="下一頁" PrevPageText="上一頁" />
           </asp:DataGrid>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DTMW_Priority.jpg" Style="z-index: 99;
               left: 0px; position: absolute; top: 8px" />
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 115; left: 480px; position: absolute; top: 40px"
               Width="145px"></asp:TextBox>
           <asp:DropDownList ID="DPriority" runat="server" BackColor="Yellow" Style="z-index: 129;
               left: 96px; position: absolute; top: 104px" Width="100px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem Value="0">HIGH</asp:ListItem>
               <asp:ListItem Value="1">MIDDLE</asp:ListItem>
               <asp:ListItem Value="2">LOW</asp:ListItem>
           </asp:DropDownList>
           <asp:DropDownList ID="DForm" runat="server" BackColor="Yellow" Style="z-index: 129;
               left: 96px; position: absolute; top: 32px" Width="288px">
           </asp:DropDownList>
           <asp:Button ID="Go" runat="server" BackColor="Aqua" ForeColor="Blue" Height="24px"
               Style="z-index: 121; left: 400px; position: absolute; top: 104px" Text="Go" Width="40px" />
           <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 116; left: 480px; position: absolute; top: 72px"
               Width="144px"></asp:TextBox>
           <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 117; left: 96px; position: absolute;
               top: 72px" Width="144px"></asp:TextBox>
  	 
      </div>
    </form>
</body>
</html>
