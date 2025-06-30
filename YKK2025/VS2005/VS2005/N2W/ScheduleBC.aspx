<%@ Page Language="vb" AutoEventWireup="false" Inherits="ScheduleBC" aspCompat="True" EnableEventValidation = "false"  CodeFile="ScheduleBC.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>出差及清算系統</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Label ID="Label1" runat="server" Style="z-index: 104; left: 248px; position: absolute;
               top: 16px" Text="出差" Font-Size="XX-Large"></asp:Label>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt" ShowFooter="True"
               Style="z-index: 102; left: 16px; position: absolute; top: 72px" Width="552px">
               <Columns>
                   <asp:BoundField DataField="StepName" HeaderText="工程" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="allcase" HeaderText="待處理件數" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="allcaseURL" HeaderText="待處理連結" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="delaycase" HeaderText="延遲件數" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="Step" HeaderText="Step" />
               </Columns>
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC"  />
           </asp:GridView>
           &nbsp;&nbsp;
           <asp:HyperLink ID="LFun02" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 102; left: 24px; position: absolute;
               top: 48px" Target="_self" Width="224px">出差調閱資料</asp:HyperLink>
           &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LFun01" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 102; left: 24px; position: absolute; top: 16px" Target="_self"
               Width="224px">出差申請新委託</asp:HyperLink>
           <asp:Label ID="Label2" runat="server" Font-Size="XX-Large" Style="z-index: 104; left: 240px;
               position: absolute; top: 408px" Text="清算"></asp:Label>
           <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt" ShowFooter="True"
               Style="z-index: 102; left: 16px; position: absolute; top: 464px" Width="544px">
               <Columns>
                   <asp:BoundField DataField="StepName" HeaderText="工程" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="allcase" HeaderText="待處理件數" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="allcaseURL" HeaderText="待處理連結" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="delaycase" HeaderText="延遲件數" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="Step" HeaderText="Step" />
               </Columns>
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC"  />
           </asp:GridView>
           <asp:HyperLink ID="LFun04" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 102; left: 16px; position: absolute; top: 432px" Target="_self"
               Width="224px">清算調閱資料</asp:HyperLink>
           <asp:HyperLink ID="LFun03" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 102; left: 16px; position: absolute; top: 400px" Target="_self"
               Width="224px">清算申請新委託</asp:HyperLink>
           &nbsp;
           <asp:Panel ID="Panel1" runat="server" Height="528px" ScrollBars="Vertical" Style="z-index: 105;
               left: 648px; position: absolute; top: 136px" Width="952px">
               <asp:GridView ID="GridView3" runat="server" BorderColor="#CC9966" BorderStyle="Groove"
                   BorderWidth="1px" CellPadding="4" Font-Size="9pt" Style="z-index: 103; left: 8px;
                   position: absolute; top: 8px" Width="672px">
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView>
           </asp:Panel>
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/BusinessTripChecklist_01.jpg"
               Style="z-index: 99; left: 632px; position: absolute; top: 8px" />
           <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="~/Images/msexcel.gif"
               Style="z-index: 100; left: 1400px; position: absolute; top: 96px" Width="21px" />
           <asp:TextBox ID="DObject" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 784px; position: absolute;
               top: 96px" Width="176px"></asp:TextBox>
           <asp:TextBox ID="DLocation" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 1104px; position: absolute;
               top: 96px" Width="176px"></asp:TextBox>
           <asp:TextBox ID="DNO" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 113; left: 1104px; position: absolute; top: 32px"
               Width="176px"></asp:TextBox>
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 113; left: 784px; position: absolute; top: 64px"
               Width="176px"></asp:TextBox>
           <asp:TextBox ID="DASDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 1104px; position: absolute;
               top: 64px" Width="64px"></asp:TextBox>
           <asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 160;
               left: 1200px; position: absolute; top: 64px; text-align: center" Text="~" Width="6px"></asp:Label>
           <asp:TextBox ID="DAEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 110; left: 1224px; position: absolute;
               top: 64px" Width="64px"></asp:TextBox>
           <asp:DropDownList ID="DSTS" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
               Style="z-index: 255; left: 784px; position: absolute; top: 32px" Width="176px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem Value="未清算">未清算</asp:ListItem>
               <asp:ListItem Value="清算中">清算中</asp:ListItem>
               <asp:ListItem Value="清算完畢">清算完畢</asp:ListItem>
           </asp:DropDownList>
           <asp:Button ID="Go" runat="server" BackColor="Aqua" ForeColor="Blue" Height="24px"
               Style="z-index: 121; left: 1352px; position: absolute; top: 96px" Text="Go" Width="40px" />
           <input id="BSDate" runat="server" style="z-index: 148; left: 1168px; width: 24px;
               position: absolute; top: 64px" type="button" value="..." visible="true" />
           <input id="BEDate" runat="server" style="z-index: 150; left: 1288px; width: 24px;
               position: absolute; top: 64px" type="button" value="..." visible="true" />
           <asp:HyperLink ID="HyperLink1" runat="server" Height="40px" ImageUrl="~/images/Ratejpg.jpg"
               NavigateUrl="~/RateList.aspx" Style="z-index: 139; left: 472px; position: absolute;
               top: 400px" Target="_blank" Width="1px" BorderStyle="Groove">HyperLink</asp:HyperLink>
       </div>
    </form>
</body>
</html>
