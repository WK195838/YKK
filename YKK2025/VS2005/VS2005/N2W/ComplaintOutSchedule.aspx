<%@ Page Language="vb" AutoEventWireup="false" Inherits="ComplaintOutSchedule" aspCompat="True" EnableEventValidation = "false"  CodeFile="ComplaintOutSchedule.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>客訴處理系統</title>
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
           <asp:Label ID="Label1" runat="server" Style="z-index: 104; left: 224px; position: absolute;
               top: 24px" Text="客訴處理系統" Font-Size="XX-Large" Width="296px"></asp:Label>
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LFun02" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               NavigateUrl="DiscountSheet_01.aspx" Style="z-index: 102; left: 856px; position: absolute;
               top: 112px" Target="_self" Width="224px">調閱資料</asp:HyperLink>
           <asp:HyperLink ID="LFun01" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               NavigateUrl="ISMSSheet_01.aspx.aspx" Style="z-index: 102; left: 856px; position: absolute;
               top: 88px" Target="_self" Width="224px">新委託</asp:HyperLink><asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt" ShowFooter="True"
               Style="z-index: 102; left: 24px; position: absolute; top: 408px" Width="790px" HorizontalAlign="Center">
                   <Columns>
                       <asp:BoundField DataField="StepName" HeaderText="工程" >
                       </asp:BoundField>
                       <asp:BoundField DataField="Data" HeaderText="相關附件" Visible="False" >
                       </asp:BoundField>
                       <asp:BoundField DataField="AllCase" HeaderText="待處理件數" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="allcaseURL" HeaderText="待處理連結" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="DelayCase" HeaderText="延遲件數" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Step" HeaderText="Step" />
                       <asp:BoundField DataField="Leadtime" HeaderText="客訴管理天數" >
                           <ItemStyle Width="80px" HorizontalAlign="Center" />
                       </asp:BoundField>
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView><asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt" ShowFooter="True"
               Style="z-index: 102; left: 24px; position: absolute; top: 88px" Width="790px" HorizontalAlign="Center">
                   <Columns>
                       <asp:BoundField DataField="StepName" HeaderText="工程" >
                       </asp:BoundField>
                       <asp:BoundField DataField="Data" HeaderText="相關附件" Visible="False" >
                       </asp:BoundField>
                       <asp:BoundField DataField="AllCase" HeaderText="待處理件數" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="allcaseURL" HeaderText="待處理連結" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="DelayCase" HeaderText="延遲件數" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Step" HeaderText="Step" />
                       <asp:BoundField DataField="Leadtime" HeaderText="客訴管理天數" >
                           <ItemStyle Width="80px" HorizontalAlign="Center" />
                       </asp:BoundField>
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView><asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt"
               Style="z-index: 102; left: 24px; position: absolute; top: 576px" Width="792px" HorizontalAlign="Center">
                   <Columns>
                       <asp:BoundField DataField="StepName" HeaderText="工程" />
                       <asp:BoundField DataField="Data" HeaderText="相關附件" Visible="False" />
                       <asp:BoundField DataField="AllCase" HeaderText="完成件數" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="allcaseURL" HeaderText="完成連結" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="DelayCase" HeaderText="延遲件數" Visible="False" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" Visible="False" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Step" HeaderText="Step" />
                       <asp:BoundField DataField="Leadtime" HeaderText="客訴管理天數" Visible="False" >
                           <ItemStyle Width="80px" HorizontalAlign="Center" />
                       </asp:BoundField>
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView><asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt"
               Style="z-index: 102; left: 24px; position: absolute; top: 664px" Width="792px" HorizontalAlign="Center">
                   <Columns>
                       <asp:BoundField DataField="StepName" HeaderText="工程" />
                       <asp:BoundField DataField="Data" HeaderText="相關附件" Visible="False" />
                       <asp:BoundField DataField="AllCase" HeaderText="完成件數（分析依賴）" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="allcaseURL" HeaderText="完成連結" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="DelayCase" HeaderText="延遲件數" Visible="False" >
                           <FooterStyle HorizontalAlign="Center" />
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" Visible="False" >
                           <ItemStyle HorizontalAlign="Center" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Step" HeaderText="Step" />
                       <asp:BoundField DataField="Leadtime" HeaderText="客訴管理天數" Visible="False" >
                           <ItemStyle Width="80px" HorizontalAlign="Center" />
                       </asp:BoundField>
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView>
           <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="12pt"
               Style="z-index: 102; left: 24px; position: absolute; top: 736px" Width="792px" HorizontalAlign="Center">
               <Columns>
                   <asp:BoundField DataField="StepName" HeaderText="工程" />
                   <asp:BoundField DataField="Data" HeaderText="相關附件" Visible="False" />
                   <asp:BoundField DataField="AllCase" HeaderText="完成件數（物料抱怨）" >
                       <FooterStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="allcaseURL" HeaderText="完成連結" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DelayCase" HeaderText="延遲件數" Visible="False" >
                       <FooterStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="delaycaseURL" HeaderText="延遲連結" Visible="False" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="Step" HeaderText="Step" />
                   <asp:BoundField DataField="Leadtime" HeaderText="客訴管理天數" Visible="False" >
                       <ItemStyle Width="80px" HorizontalAlign="Center" />
                   </asp:BoundField>
               </Columns>
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:GridView>
       </div>
    </form>
</body>
</html>
