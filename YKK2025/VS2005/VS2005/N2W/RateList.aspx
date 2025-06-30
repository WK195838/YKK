<%@ Page Language="vb" AutoEventWireup="false" Inherits="RateList" aspCompat="True" EnableEventValidation = "false"  CodeFile="RateList.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>外幣匯率</title>
 <script type="text/javascript">  
 
 	//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			

</script>  
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" CellPadding="4" Font-Size="9pt" Style="z-index: 100; left: 8px; position: absolute; top: 168px"
               Width="790px" ForeColor="#333333" GridLines="None" AutoGenerateColumns="False">
               <FooterStyle BackColor="#5D7B9D" ForeColor="White" Font-Bold="True" />
               <HeaderStyle BackColor="#E0E0E0" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="False" />
               <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
               <PagerStyle ForeColor="White" HorizontalAlign="Center" BackColor="#284775" />
               <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" HorizontalAlign="Center" />
               <EditRowStyle BackColor="#999999" />
               <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            
               <EmptyDataTemplate>
                   
               </EmptyDataTemplate>
               <Columns>
                   <asp:BoundField DataField="seqno" >
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:TemplateField HeaderText="國家名&lt;/br&gt;(Country)">
                       <ItemTemplate>
                             <asp:Label ID="Label1" runat="server"
            Text='<%# Bind("country") %>'></asp:Label>
                       </ItemTemplate>
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                     <asp:TemplateField HeaderText="幣別名&lt;/br&gt;(Currency)">
                       <ItemTemplate>
                             <asp:Label ID="Label2" runat="server"
            Text='<%# Bind("currency") %>'></asp:Label>
                       </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                     <asp:TemplateField HeaderText="幣別代號&lt;/br&gt;(Currencycode)">
                       <ItemTemplate>
                             <asp:Label ID="Label3" runat="server"
            Text='<%# Bind("currencycode") %>'></asp:Label>
                       </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                    <asp:TemplateField HeaderText="年&lt;/br&gt;(Year)">
                       <ItemTemplate>
                             <asp:Label ID="Label4" runat="server"
            Text='<%# Bind("syear") %>'></asp:Label>
                       </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                     <asp:TemplateField HeaderText="月&lt;/br&gt;(Month)">
                       <ItemTemplate>
                             <asp:Label ID="Label5" runat="server"
            Text='<%# Bind("smonth") %>'></asp:Label>
                       </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                     <asp:TemplateField HeaderText="日&lt;/br&gt;(Period)">
                       <ItemTemplate>
                             <asp:Label ID="Label6" runat="server"
            Text='<%# Bind("sday") %>'></asp:Label>
                       </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                    <asp:TemplateField HeaderText="平均匯率&lt;/br&gt;(Average)">
                       <ItemTemplate>
                             <asp:Label ID="Label6" runat="server"
            Text='<%# Bind("averagerate") %>'></asp:Label>
                       </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" />
                   </asp:TemplateField>
                   <asp:BoundField DataField="StrD" HeaderText="StrD" />
                   <asp:BoundField DataField="EndD" HeaderText="EndD" />
                   <asp:BoundField DataField="Data1" HeaderText="Data1" />
               </Columns>
        
          
  	 
           </asp:GridView>
        
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/RateSelect.jpg" Style="z-index: 101;
               left: 8px; position: absolute; top: 16px" />
           <asp:DropDownList ID="DSYear" runat="server" Style="z-index: 102; left: 264px; position: absolute; top: 72px"
               Width="72px" AutoPostBack="True" BackColor="Yellow">
           </asp:DropDownList>
           <asp:DropDownList ID="DSMonth" runat="server" Style="z-index: 103; left: 408px; position: absolute; top: 72px"
               Width="88px" AutoPostBack="True" BackColor="Yellow">
           </asp:DropDownList>
           <asp:DropDownList ID="DSDay" runat="server" Style="z-index: 104; left: 592px; position: absolute; top: 72px"
               Width="96px" AutoPostBack="True" BackColor="Yellow">
           </asp:DropDownList>
           <asp:DropDownList ID="DCurrency" runat="server" Style="z-index: 105; left: 264px; position: absolute; top: 24px"
               Width="320px" BackColor="Yellow">
           </asp:DropDownList>
       
           <asp:Button ID="Button1" runat="server" BackColor="#FF8000" ForeColor="White" Style="z-index: 108;
               left: 592px; position: absolute; top: 136px" Text="查詢" Width="72px" />
           <asp:Button ID="Button2" runat="server" BackColor="#FF8000" ForeColor="White" Style="z-index: 108;
               left: 680px; position: absolute; top: 136px" Text="匯出" Width="72px" />
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;
  	 
      </div>
    </form>
</body>
</html>
