<%@ Page Language="vb" AutoEventWireup="false" Inherits="QASheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="QASheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>品質分析依賴書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
   function GetCustomer()
{
        
    window.open('CustomerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 
            


   function GetBuyer()
{
        
    window.open('BuyerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
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
               <img src="images/QAHead_03.jpg" style="z-index: 1; left: 16px; position: absolute;
                   top: 0px" />
               <asp:HyperLink ID="LModNo" runat="server" Style="z-index: 115; left: 560px; position: absolute;
                   top: 40px" Target="_blank" Visible="False">有修改申請</asp:HyperLink>
               <asp:RadioButtonList ID="DLocation" runat="server" Style="z-index: 101; left: 560px;
                   position: absolute; top: 536px">
               </asp:RadioButtonList><asp:RadioButtonList ID="DSample" runat="server" Style="z-index: 101; left: 560px;
                   position: absolute; top: 472px">
               </asp:RadioButtonList><asp:RadioButtonList ID="DMaterial" runat="server" Style="z-index: 101; left: 560px;
                   position: absolute; top: 504px">
               </asp:RadioButtonList>
               &nbsp;
               &nbsp;&nbsp;
               <asp:CheckBoxList ID="DCheckItem" runat="server" RepeatDirection="Horizontal" Style="z-index: 102;
                   left: 155px; position: absolute; top: 240px">
               </asp:CheckBoxList>
               &nbsp;
                  <asp:TextBox ID="DUpdate" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 101; left: 1208px; position: absolute; top: 264px"
                   Width="142px"></asp:TextBox>
                   
               <asp:Button ID="DAttachFile2" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 160px; position: absolute; top: 600px" Text="開啟附件" Width="72px" />
               &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
               <asp:CheckBoxList ID="DRPReport" runat="server" RepeatDirection="Horizontal"
                   Style="z-index: 102; left: 560px; position: absolute; top: 440px">
               </asp:CheckBoxList>
               &nbsp;
               <asp:CheckBoxList ID="DCheckType" runat="server" RepeatDirection="Horizontal"
                   Style="z-index: 102; left: 155px; position: absolute; top: 376px">
               </asp:CheckBoxList>
               &nbsp;
               <asp:TextBox ID="DQCDate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px; position: absolute;
                   top: 576px; text-align: left" Width="96px"></asp:TextBox>
               &nbsp;&nbsp;
               <asp:Button ID="DAttachFile1" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 560px; position: absolute; top: 408px" Text="開啟附件" Width="72px" />
               &nbsp; &nbsp;
               <asp:TextBox ID="DFinishDate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px; position: absolute;
                   top: 376px; text-align: left" Width="96px"></asp:TextBox>
               &nbsp; &nbsp;
               <asp:TextBox ID="DADate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px; position: absolute;
                   top: 112px; text-align: left" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DBuyer" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 110; left: 560px; position: absolute;
                   top: 344px" Width="256px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
                   BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 111; left: 560px;
                   position: absolute; top: 312px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DCustomer" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 112; left: 560px; position: absolute;
                   top: 272px" Width="256px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DCustomerCode" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 113; left: 560px; position: absolute;
                   top: 240px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 112px; text-align: left" Width="144px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DQCNo" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 568px; text-align: left" Width="240px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DItemName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 176px; text-align: left" Width="656px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DBackGround" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 208px; text-align: left" Width="656px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 80px; text-align: left" Width="240px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 560px; position: absolute; top: 80px" Width="120px">DNo</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 312px;
                   position: absolute; top: 112px; text-align: left" Width="88px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;&nbsp;
               <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
                   Style="z-index: 137; left: 32px; position: absolute; top: 952px" Text="核定履歷"></asp:Label>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                   Style="z-index: 136; left: 32px; position: absolute; top: 976px" Width="780px">
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
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DSubject" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 160px; position: absolute;
               top: 144px; text-align: left" Width="656px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DTNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 132; left: 952px; position: absolute;
               top: 96px" Width="190px"></asp:TextBox>
           <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="208px" ScrollBars="Auto"
               Style="left: 16px; position: absolute; top: 696px" Width="1712px">
               &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
               <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="Solid"
                   BorderWidth="1px" CellPadding="3" DataKeyNames="Unique_ID" Font-Size="9pt" ForeColor="Black"
                   GridLines="Vertical" PageSize="100" Style="z-index: 100; left: 8px; position: absolute;
                   top: 8px" Width="1384px">
                   <Columns>
                       <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="30px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="SeqNo" HeaderText="No" />
                       <asp:BoundField DataField="Type" HeaderText="類別">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Supplier" HeaderText="廠商">
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Code" HeaderText="CODE">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Size" HeaderText="SIZE">
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Family" HeaderText="FAMILY">
                           <HeaderStyle Height="20px" />
                           <ItemStyle HorizontalAlign="Left" Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Body" HeaderText="BODY">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Puller" HeaderText="PULLER">
                           <HeaderStyle Height="20px" />
                           <ItemStyle HorizontalAlign="Left" Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField HeaderText="COLOR" DataField="COLOR" />
                       <asp:BoundField DataField="Finish" HeaderText="FINISH">
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="ApproveNo" HeaderText="共用核可卡">
                           <ItemStyle HorizontalAlign="Left" Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="OldToNewNo" HeaderText="舊換新NO" >
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="OldToNew" HeaderText="舊換新" >
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Remark" HeaderText="備註">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Result1" HeaderText="Test Result" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Result2" HeaderText="覆判結果" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField HeaderText="QC備註" >
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="OutTestNo" HeaderText="外測連結">
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Formula" HeaderText="色粉配方連結">
                           <ItemStyle Width="120px" />
                       </asp:BoundField>
                   </Columns>
                   <FooterStyle BackColor="#CCCCCC" />
                   <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                   <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                   <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                       VerticalAlign="Middle" />
                   <AlternatingRowStyle BackColor="#CCCCCC" />
               </asp:GridView>
               &nbsp;
           </asp:Panel>
           <asp:TextBox ID="DCODE" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 952px; position: absolute;
               top: 272px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="DPaletNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 952px; position: absolute;
               top: 232px" Width="190px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DCount" runat="server" BackColor="Yellow" BorderStyle="Groove"
               Font-Size="Larger" ForeColor="Blue" Height="27px" Style="z-index: 102; left: 848px;
               position: absolute; top: 664px" Width="94px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 138; left: 1144px; position: absolute; top: 336px"
               Width="190px"></asp:TextBox>
           <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1152px; position: absolute;
               top: 200px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="D3" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1144px; position: absolute;
               top: 248px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="DHaveData" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1144px; position: absolute;
               top: 360px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="CheckTypeCount" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1000px; position: absolute;
               top: 480px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="RPReportCount" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1120px; position: absolute;
               top: 472px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="SampleCount" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 896px; position: absolute;
               top: 528px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="MaterialCount" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1000px; position: absolute;
               top: 520px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="LocationCount" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 1120px; position: absolute;
               top: 512px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="CheckItemCount" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 138; left: 896px; position: absolute;
               top: 480px" Width="64px"></asp:TextBox>
  	 
      </div>
    </form>
</body>
</html>
