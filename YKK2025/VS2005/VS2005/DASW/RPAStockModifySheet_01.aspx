<%@ Page Language="vb" AutoEventWireup="false" Inherits="RPAStockModifySheet_01" aspCompat="True" EnableEventValidation = "false"  CodeFile="RPAStockModifySheet_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<script runat="server">

       
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>RPA在庫修正申請表</title>
    	<script language="javascript" type="text/javascript">
//			var wPop;
//			--Calendar------------------------------------
//			function calendarPicker(strField)
//			{
//				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
//			}

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
function IMG1_onclick() {

}

		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="Images/RPAStockModify_02.jpg" style="z-index: 99; left: 8px; position: absolute;
                   top: 16px" id="IMG1" onclick="return IMG1_onclick()" />
               <asp:HyperLink ID="LRPAFILE" runat="server" Style="z-index: 138; left: 616px; position: absolute;
                   top: 304px" Visible="False">RPA回傳檔</asp:HyperLink>
               <asp:TextBox ID="DRPAComment" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 121; left: 136px;
                   position: absolute; top: 296px; text-align: left" Width="472px"></asp:TextBox>
               <asp:DropDownList ID="DRPAREASON" runat="server" BackColor="Yellow"
                   ForeColor="Black" Height="24px" Style="z-index: 105; left: 136px; position: absolute;
                   top: 200px" Width="560px">
               </asp:DropDownList>
               &nbsp;&nbsp;&nbsp;&nbsp;
                <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 118; left: 136px; position: absolute;
                   top: 96px; text-align: left" Width="220px"></asp:TextBox>&nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 119;
                   left: 480px; position: absolute; top: 96px"></asp:TextBox>
               <asp:TextBox ID="DDepID" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 126; left: 136px;
                   position: absolute; top: 128px; text-align: left" Width="80px"></asp:TextBox>
               &nbsp;               
               <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 126; left: 232px;
                   position: absolute; top: 128px; text-align: left" Width="128px"></asp:TextBox>
               &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;
               &nbsp;
               <asp:TextBox ID="DAppName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="24px" MaxLength="7" Style="z-index: 121; left: 472px;
                   position: absolute; top: 128px; text-align: left" Width="224px"></asp:TextBox>
               <asp:TextBox ID="DAMOUNT" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="24px" MaxLength="7" Style="z-index: 121; left: 472px; position: absolute;
                   top: 160px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;
               &nbsp;
               &nbsp;    
               &nbsp;&nbsp;              
               
               <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="50px" MaxLength="255" Style="z-index: 101; left: 136px;
                   position: absolute; top: 232px; text-align: left" TextMode="MultiLine" Width="560px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
               &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
              
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
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
               &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 141; left: 32px;
                   position: absolute; top: 424px" visible="false" />
               &nbsp;
               <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
                   Style="z-index: 122; left: 264px; position: absolute; top: 432px" Visible="False"
                   Width="352px"></asp:TextBox>
               <asp:DropDownList ID="DReasonCode" runat="server" BackColor="Yellow" Height="20px"
                   Style="z-index: 123; left: 192px; position: absolute; top: 440px" Visible="False"
                   Width="64px">
                   <asp:ListItem>01</asp:ListItem>
                   <asp:ListItem>02</asp:ListItem>
               </asp:DropDownList>
               <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 124; left: 192px; position: absolute; top: 464px"
                   Visible="False" Width="424px"></asp:TextBox>
               <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
                   Style="z-index: 125; left: 24px; position: absolute; top: 592px" Text="核定履歷"></asp:Label>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                   Style="z-index: 126; left: 16px; position: absolute; top: 616px" Width="780px">
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
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 132; left: -500px; position: absolute;
                   top: 610px; text-align: left">AAA</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 133; left: 360px; position: absolute; top: 544px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 134; left: 448px; position: absolute; top: 544px" Text="NG2"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 135; left: 544px; position: absolute; top: 544px" Text="NG1"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 136; left: 632px; position: absolute; top: 544px" Text="OK"
                   UseSubmitBehavior="false" Width="80px" OnClick="BOK_Click" />
           </div>
       
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LAttachfile" runat="server" Style="z-index: 138; left: 136px;
               position: absolute; top: 168px">附檔</asp:HyperLink>
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="56px" Style="z-index: 132; left: 80px; position: absolute; top: 360px"
               TextMode="MultiLine" Width="536px" Text="OK"></asp:TextBox>
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 1;
               left: 32px; position: absolute; top: 352px" />
           &nbsp;&nbsp;&nbsp;
           <input id="DDISPOSALFILE1" runat="server" name="DISPOSALFILE1" style="z-index: 293;
               left: 1064px; width: 270px; position: absolute; top: 248px; height: 26px; background-color: #ffff00"
               type="file" visible="false" />
           <asp:Button ID="DUpload" runat="server" CausesValidation="False" Style="z-index: 279;
               left: 1064px; position: absolute; top: 296px" Text="預覽" Width="47px" Visible="False" />
           <asp:Button ID="DInsert" runat="server" Style="z-index: 280; left: 1120px; position: absolute;
               top: 296px" Text="上傳" UseSubmitBehavior="False" Width="47px" Visible="False" />
           <asp:RadioButtonList ID="rbHDR" runat="server" Style="z-index: 281; left: 1056px;
               position: absolute; top: 168px" Visible="False">
               <asp:ListItem Selected="True" Text="Yes" Value="Yes"></asp:ListItem>
               <asp:ListItem Text="No" Value="No"></asp:ListItem>
           </asp:RadioButtonList>
           <asp:GridView ID="GridView1" runat="server" Style="z-index: 282; left: 1072px; position: absolute;
               top: 360px" Visible="False">
           </asp:GridView>
           <asp:FileUpload ID="DFileUpload1" runat="server" BackColor="Yellow" Height="22px"
               Style="z-index: 278; left: 136px; position: absolute; top: 168px" Width="224px" />
           </div>
    </form>
</body>
</html>
