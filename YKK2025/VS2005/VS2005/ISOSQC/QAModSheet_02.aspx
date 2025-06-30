<%@ Page Language="vb" AutoEventWireup="false" Inherits="QAModSheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="QAModSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>品質分析資料修改依賴書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------		


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
        	<FONT face="新細明體"></FONT>
           <br />
           <br />
           <br />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
            <img src="images/ModQAHead_01.jpg" style="z-index: 2; left: 16px; position: absolute;
                   top: 0px" />
               <asp:TextBox ID="ModContent" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="100px" Style="z-index: 126; left: 140px; position: absolute;
                   top: 257px; text-align: left" Width="540px" TextMode="MultiLine"></asp:TextBox>
               <asp:TextBox ID="IContent" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="80px" Style="z-index: 126; left: 140px; position: absolute;
                   top: 409px; text-align: left" Width="540px" TextMode="MultiLine" Enabled="false"></asp:TextBox>
               <img src="images/QAHead_01.jpg" style="z-index: 1; left: 16px; position: absolute;
                   top: 530px" id="IMG1" />
               <asp:RadioButtonList ID="DLocation" runat="server" Style="z-index: 101; left: 560px;
                   position: absolute; top: 1066px">
               </asp:RadioButtonList><asp:RadioButtonList ID="DSample" runat="server" Style="z-index: 101; left: 560px;
                   position: absolute; top: 1002px">
               </asp:RadioButtonList><asp:RadioButtonList ID="DMaterial" runat="server" Style="z-index: 101; left: 560px;
                   position: absolute; top: 1034px">
               </asp:RadioButtonList>
               &nbsp;
               &nbsp;&nbsp;
               <asp:CheckBoxList ID="DCheckItem" runat="server" RepeatDirection="Horizontal" Style="z-index: 102;
                   left: 155px; position: absolute; top: 770px">
               </asp:CheckBoxList>
               &nbsp;
                  <asp:TextBox ID="DUpdate" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 101; left: 1208px; position: absolute; top: 264px"
                   Width="142px"></asp:TextBox>
                   
               <asp:Button ID="DAttachFile2" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 160px; position: absolute; top: 1130px" Text="開啟附件" Visible="false" Width="72px" />
               &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
               <asp:CheckBoxList ID="DRPReport" runat="server" RepeatDirection="Horizontal"
                   Style="z-index: 102; left: 560px; position: absolute; top: 970px">
               </asp:CheckBoxList>
               &nbsp;
               <asp:CheckBoxList ID="DCheckType" runat="server" RepeatDirection="Horizontal"
                   Style="z-index: 102; left: 155px; position: absolute; top: 906px">
               </asp:CheckBoxList>
               <asp:TextBox ID="DQCDate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px; position: absolute;
                   top: 1106px; text-align: left" Width="96px"></asp:TextBox>
               &nbsp;&nbsp;
               <asp:Button ID="DAttachFile1" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 560px; position: absolute; top: 938px" Text="開啟附件" Visible="false" Width="72px" />
               &nbsp;
               <asp:TextBox ID="DFinishDate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px; position: absolute;
                   top: 906px; text-align: left" Width="96px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DADate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px; position: absolute;
                   top: 642px; text-align: left" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DBuyer" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 110; left: 560px; position: absolute;
                   top: 874px" Width="256px"></asp:TextBox>
               <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
                   BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 111; left: 560px;
                   position: absolute; top: 842px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DCustomer" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 112; left: 560px; position: absolute;
                   top: 802px" Width="256px"></asp:TextBox>
               <asp:TextBox ID="DCustomerCode" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 113; left: 560px; position: absolute;
                   top: 770px" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 140px; position: absolute;
                   top: 111px; text-align: left" Width="210px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DQCNo" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 1098px; text-align: left" Width="240px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DItemName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 706px; text-align: left" Width="656px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DBackGround" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 738px; text-align: left" Width="656px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" Style="z-index: 126; left: 140px; position: absolute;
                   top: 76px; text-align: left" Width="210px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 475px; position: absolute; top: 76px" Width="120px" Text="DNo"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="ModNo" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 127; left: 560px; position: absolute;
                   top: 607px; text-align: left" Width="96px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 475px;
                   position: absolute; top: 108px; text-align: left" Width="180px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="ModDate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 607px; text-align: left" Width="240px"></asp:TextBox>
               <asp:TextBox ID="ModDepname" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
                   top: 642px; text-align: left" Width="144px"></asp:TextBox>
               <asp:DropDownList ID="DCheckUser" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 140px; position: absolute;
                   top: 145px; text-align: left" Width="210px"></asp:DropDownList>
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
                   Style="z-index: 137; left: 14px; position: absolute; top: 1487px" Text="核定履歷"></asp:Label>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                   Style="z-index: 136; left: 13px; position: absolute; top: 1509px" Width="780px">
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
               top: 674px; text-align: left" Width="656px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DTNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 132; left: 952px; position: absolute;
               top: 96px" Width="190px"></asp:TextBox>
           <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="248px" ScrollBars="Auto"
               Style="left: 24px; position: absolute; top: 1218px" Width="1216px">
               &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
               <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" BorderStyle="Solid"
                   BorderWidth="1px" CellPadding="3" DataKeyNames="Unique_ID" Font-Size="9pt" ForeColor="Black"
                   GridLines="Vertical" PageSize="100" Style="z-index: 100; left: 9px; position: absolute;
                   top: 40px">
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
                       <asp:BoundField DataField="COLOR" HeaderText="COLOR" />
                       <asp:BoundField DataField="Finish" HeaderText="FINISH">
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="ApproveNo" HeaderText="共用核可卡">
                           <ItemStyle HorizontalAlign="Right" />
                       </asp:BoundField>
                       <asp:BoundField DataField="OldToNewNo" HeaderText="舊換新NO" />
                       <asp:BoundField DataField="OldToNew" HeaderText="舊換新" />
                       <asp:BoundField DataField="Remark" HeaderText="備註">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Element" HeaderText="待判(element)											">
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Result1" HeaderText="Test Result" />
                       <asp:BoundField DataField="Result2" HeaderText="覆判結果" />
                       <asp:BoundField DataField="QCRemark" HeaderText="QC備註" />
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
               position: absolute; top: 1178px" Width="94px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp;
           <asp:Button ID="BAdd" runat="server" Height="30px" Style="z-index: 132; left: 32px;
               position: absolute; top: 1178px" Text=" 細項登錄" UseSubmitBehavior="false" Width="152px" BackColor="#80FF80" BorderStyle="Outset" Font-Size="Medium" />
           <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 138; left: 1144px; position: absolute; top: 336px"
               Width="190px"></asp:TextBox>
           &nbsp;
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
           <asp:TextBox ID="ModName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 312px; position: absolute;
               top: 642px; text-align: left" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DNO1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 475px; position: absolute;
               top: 144px; text-align: left" Width="180px"></asp:TextBox>
               <asp:hyperlink id="LSheet1" style="Z-INDEX: 142; LEFT: 570px; POSITION: absolute; TOP: 143px" runat="server"
					Width="65px" Height="8px" Font-Size="12pt" NavigateUrl="QASheet_02.aspx" Target="_blank">連結表單</asp:hyperlink>
           <asp:TextBox ID="ModReason" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="60px" Style="z-index: 126; left: 140px; position: absolute;
               top: 177px; text-align: left" Width="540px" TextMode="MultiLine"></asp:TextBox>
           <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 138; left: 1152px; position: absolute; top: 200px"
               Width="190px"></asp:TextBox>

  	 
      </div>
    </form>
</body>
</html>
