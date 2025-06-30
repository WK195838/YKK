<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CloseAccountSheet_012.aspx.vb" Inherits="CloseAccountSheet_012" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>清算申請</title>

    <script language="javascript" type="text/javascript">
		function calendarPicker(strField)
		{
			window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
		}
        			
       function GetEMP()
       {
            if (document.getElementById("DDivision2").value == "")
            {
                var Div = document.getElementById("DDivision1").value;
            }
            else
            {
                var Div = document.getElementById("DDivision2").value;
            }
           
            window.open('EMPList.aspx?pKey=' + Div,'EMPPopup','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
       }
        			
       function GetDivision()
       {
            if (document.getElementById("DDivision3").value == "")
            {
                var Div = document.getElementById("DDivision1").value;
            }
            else
            {
                var Div = document.getElementById("DDivision3").value;
            }
            
            window.open('DivisionList.aspx?pKey=' + Div,'DivisionPopup','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
       }

       function GetPassport(strField)
	   {
		    window.open('PassportList.aspx?field=' + strField,'Popup','width=600,height=350,top=10,resizable=yes');
	   }

       function GetTarget(strField)
       {
		    window.open('TargetList.aspx?field=' + strField,'Popup','width=600,height=350,top=10,resizable=yes');
       }
        			
       function AddTrip()
       {
            window.open('BusinessTripList.aspx','_blank');
       }
       
          			
       function AddExpense()
       {
            window.open('ExpenseList.aspx','_blank');
       }
       		
       		
  
	</script>
</head>


<body>
  	<form id="Form1" runat="server">
         <div>
            <!-- -->
            <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
            <!-- ++  底圖                                                                                ++ -->
            <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
             &nbsp;<!-- --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  欄位                                                                                ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- --><!-- ++  申請日/表單NO ++ -->
            <asp:TextBox ID="DDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 100; left: 152px; position: absolute; top: 112px" Width="152px" ReadOnly="true"></asp:TextBox>
            <!-- -->
            <!-- ++  申請者：部門/姓名/CODE/ID ++ -->
            <asp:TextBox ID="DDivision1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 101; left: 152px; position: absolute; top: 144px" Width="80px" ReadOnly="true"></asp:TextBox>
            <asp:TextBox ID="DEmpName1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 102; left: 240px; position: absolute; top: 144px" Width="80px" ReadOnly="true"></asp:TextBox>
            <asp:TextBox ID="DDivisionCode1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 103; left: 328px; position: absolute; top: 144px" Width="80px" ReadOnly="true"></asp:TextBox>
            <asp:TextBox ID="DEmpID1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 104; left: 416px; position: absolute; top: 144px" Width="80px" ReadOnly="true"></asp:TextBox>
            <asp:TextBox ID="DJobTitle1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 105; left: 504px; position: absolute; top: 144px" Width="80px" ReadOnly="true"></asp:TextBox>
            <!-- -->
            <!-- ++  費用歸屬：部門/CODE ++ -->
            <asp:TextBox ID="DDivision3" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 106; left: 712px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
            <asp:TextBox ID="DDivisionCode3" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 107; left: 800px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
             &nbsp;&nbsp;<!-- ++   CHECKBOX / 按鈕  -->
            <!-- -->
            <!-- ++  代理人：部門/姓名/CODE/ID ++ -->
            <asp:TextBox ID="DDivision2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 108; left: 152px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
            <asp:TextBox ID="DEmpName2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 109; left: 240px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
            <asp:TextBox ID="DDivisionCode2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 110; left: 328px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
            <asp:TextBox ID="DEmpID2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 111; left: 416px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
            <asp:TextBox ID="DJobTitle2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 112; left: 504px; position: absolute; top: 176px" Width="80px"></asp:TextBox>
             &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<!-- ++   按鈕  -->
            <!-- -->
            <!-- ++  PASSPORT ++ -->
             <!-- ++   按鈕  -->
             &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- --><!-- ++  出差 ++ -->
             &nbsp;&nbsp;<!-- ++   按鈕  -->
             &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  表格 ++ -->
            <!-- ------------------------------------------------------------------------------------------ -->
            <!-- ++  追加 /  ++ -->
             &nbsp; &nbsp;&nbsp;
             <asp:Image ID="Image1" runat="server" ImageUrl="~/images/CloseAccountSheet_011.jpg"
                 Style="z-index: 99; left: 8px; position: absolute; top: 8px" />
             <asp:TextBox ID="DQCNO" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                 Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Height="24px"
                 Style="z-index: 123; left: 1080px; position: absolute; top: 176px" Width="88px"></asp:TextBox>
             <asp:Button ID="BAdd" runat="server" Style="z-index: 114;
                 left: 192px; position: absolute; top: 288px" Text="登錄" Width="88px" Visible="False" /><input id="BAddTrip" runat="server" style="z-index: 149; left: 936px; width: 24px;
                 position: absolute; top: 144px" type="button" value="..." />
             &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
             &nbsp;&nbsp;
             <asp:Button ID="BCheck" runat="server" Style="z-index: 114;
                 left: 96px; position: absolute; top: 288px" Text="檢查" Width="88px" />
             &nbsp; &nbsp;
             <asp:TextBox ID="DSumAmt" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 115; left: 1024px; position: absolute;
                 top: 296px; text-align: right" Width="152px">0</asp:TextBox>
             &nbsp; &nbsp; &nbsp;
             &nbsp; &nbsp; &nbsp; &nbsp;
             <asp:Label ID="Label1" runat="server" Style="z-index: 116; left: 256px; position: absolute;
                 top: 248px" Text="~"></asp:Label>
             <asp:HyperLink ID="LTripNo" runat="server" Style="z-index: 117; left: 976px; position: absolute;
                 top: 152px" Target="_blank" Visible="False">Link</asp:HyperLink>
             <asp:HyperLink ID="LQCNo" runat="server" Style="z-index: 118; left: 1008px;
                 position: absolute; top: 184px" Target="_blank" Visible="False">[LQCNo]</asp:HyperLink>
             &nbsp;
             <asp:TextBox ID="DASDate" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 119; left: 712px; position: absolute;
                 top: 240px" Width="88px" AutoPostBack="True"  ></asp:TextBox>
             <input id="BSDate" runat="server" style="z-index: 148; left: 808px; width: 24px;
                 position: absolute; top: 240px" type="button" value="..." visible="true" />
             <asp:TextBox ID="DAEDate" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 120; left: 840px; position: absolute;
                 top: 240px" Width="88px" AutoPostBack="True"></asp:TextBox>
             <input id="BEDate" runat="server" style="z-index: 150; left: 936px; width: 24px;
                 position: absolute; top: 240px" type="button" value="..." visible="true" />
             <asp:TextBox ID="DSDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 121; left: 152px; position: absolute;
                 top: 240px" Width="88px"></asp:TextBox>
             <asp:TextBox ID="DEDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 122; left: 280px; position: absolute;
                 top: 240px" Width="88px"></asp:TextBox>
             <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                 Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Height="24px"
                 Style="z-index: 123; left: 712px; position: absolute; top: 112px" Width="208px"></asp:TextBox>
             &nbsp; &nbsp;
             &nbsp;
             &nbsp;&nbsp;
             <asp:TextBox ID="DObject" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 125; left: 152px; position: absolute;
                 top: 208px; text-align: left" Width="432px"></asp:TextBox>
             <asp:TextBox ID="DLocation" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 126; left: 712px; position: absolute;
                 top: 208px; text-align: left" Width="456px"></asp:TextBox>
             &nbsp;
             <asp:TextBox ID="DTripNo" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
                 ForeColor="Blue" Height="18px" Style="z-index: 128; left: 712px; position: absolute;
                 top: 144px; text-align: left" Width="208px"></asp:TextBox>
             <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
                 Style="z-index: 129; left: 16px; position: absolute; top: 1000px" Text="核定履歷"></asp:Label>
             <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                 BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                 Style="z-index: 130; left: 16px; position: absolute; top: 1024px" Width="780px">
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
             <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                 Style="z-index: 131; left: 360px; position: absolute; top: 976px" Text="儲存" UseSubmitBehavior="false"
                 Width="80px" />
             <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                 Style="z-index: 132; left: 448px; position: absolute; top: 976px" Text="NG2"
                 UseSubmitBehavior="false" Width="80px" />
             <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                 Style="z-index: 133; left: 544px; position: absolute; top: 976px" Text="NG1"
                 UseSubmitBehavior="false" Width="80px" />
             <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                 Style="z-index: 134; left: 632px; position: absolute; top: 976px" Text="OK" UseSubmitBehavior="false"
                 Width="80px" />
             &nbsp; &nbsp; &nbsp; &nbsp;
             <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="312px" ScrollBars="Auto"
                 Style="left: 16px; position: absolute; top: 328px; z-index: 135;" Width="1184px">
                 &nbsp; &nbsp;&nbsp; 
                 <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="White"
                     BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" CellPadding="3" DataKeyNames="Unique_ID"
                     Font-Size="9pt" ForeColor="Black" GridLines="Vertical" PageSize="100" Style="z-index: 103;
                     left: 8px; position: absolute; top: 8px">
                     <Columns>
                         <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="100px" />
                         </asp:BoundField>
                         <asp:BoundField DataField="Type" HeaderText="類別用途">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="100px" />
                         </asp:BoundField>
                         <asp:BoundField DataField="Pay" HeaderText="支付" />
                         <asp:BoundField DataField="SDate" HeaderText="日期">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="50px" />
                         </asp:BoundField>
                         <asp:BoundField DataField="Currency" HeaderText="幣別">
                             <ItemStyle Width="100px" />
                         </asp:BoundField>
                         <asp:BoundField DataField="Money" HeaderText="金額" DataFormatString="{0:N2}">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="100px" HorizontalAlign="Right" />
                         </asp:BoundField>
                         <asp:BoundField DataField="Days" HeaderText="數量/次數">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="100px" />
                         </asp:BoundField>
                         <asp:BoundField DataField="Rate" HeaderText="匯率">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="50px" HorizontalAlign="Right" />
                         </asp:BoundField>
                         <asp:BoundField DataField="SumAmt" HeaderText="小計金額（台幣)" DataFormatString="{0:N0}" >
                             <ItemStyle HorizontalAlign="Right" />
                         </asp:BoundField>
                         <asp:BoundField DataField="ExpenseNo" HeaderText="交際費單號" />
                         <asp:BoundField DataField="Remark" HeaderText="備註">
                             <HeaderStyle Height="20px" />
                             <ItemStyle Width="200px" />
                         </asp:BoundField>
                     </Columns>
                     <FooterStyle BackColor="#CCCCCC" />
                     <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                     <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                     <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                         VerticalAlign="Middle" />
                     <AlternatingRowStyle BackColor="#CCCCCC" />
                 </asp:GridView>
                 &nbsp; &nbsp;&nbsp;
                 </asp:Panel>
             <asp:TextBox ID="D1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                 Height="24px" Style="z-index: 136; left: 1264px; position: absolute; top: 256px"
                 Width="142px"></asp:TextBox>
             &nbsp; &nbsp;&nbsp;
             <asp:TextBox ID="DTNo" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                 Height="24px" Style="z-index: 137; left: 1280px; position: absolute; top: 160px"
                 Width="142px"></asp:TextBox>
             <asp:TextBox ID="DUpdate" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                 Height="24px" Style="z-index: 138; left: 1256px; position: absolute; top: 344px"
                 Width="142px"></asp:TextBox>
             &nbsp; &nbsp; &nbsp;&nbsp;<br />
             &nbsp;
             &nbsp;&nbsp;&nbsp;<br />
             <asp:TextBox ID="DDays" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                 Height="24px" Style="z-index: 138; left: 1256px; position: absolute; top: 384px"
                 Width="142px"></asp:TextBox>
             <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                 Height="56px" Style="z-index: 132; left: 56px; position: absolute; top: 768px"
                 TextMode="MultiLine" Width="536px"></asp:TextBox>
             <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
                 left: 8px; position: absolute; top: 760px" />
             <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 193;
                 left: 8px; position: absolute; top: 840px" visible="false" />
             <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
                 Style="z-index: 132; left: 240px; position: absolute; top: 848px" Visible="False"
                 Width="352px"></asp:TextBox>
             <asp:DropDownList ID="DReasonCode" runat="server" BackColor="Yellow" Height="20px"
                 Style="z-index: 133; left: 160px; position: absolute; top: 848px" Visible="False"
                 Width="64px">
                 <asp:ListItem>01</asp:ListItem>
                 <asp:ListItem>02</asp:ListItem>
             </asp:DropDownList>
             <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                 Height="56px" Style="z-index: 134; left: 168px; position: absolute; top: 880px"
                 Visible="False" Width="424px"></asp:TextBox>
          </div>
    </form>
</body>
</html>