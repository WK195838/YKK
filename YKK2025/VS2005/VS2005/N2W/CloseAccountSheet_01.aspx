<%@ Page Language="vb" AutoEventWireup="false" Inherits="CloseAccountSheet_01" aspCompat="True" EnableEventValidation = "false"  CodeFile="CloseAccountSheet_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
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
            window.open('BusinessTripList.aspx','newwindow','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
       }
       
    
          			
       function AddExpense()
       {
            window.open('ExpenseList.aspx','newwindow','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
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
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 147; left: 12px;
                   position: absolute; top: 1034px" visible="false" />
               &nbsp;
               <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
                   Style="z-index: 100; left: 243px; position: absolute; top: 1039px" Visible="False"
                   Width="352px"></asp:TextBox>
               <asp:DropDownList ID="DReasonCode" runat="server" BackColor="Yellow" Height="20px"
                   Style="z-index: 101; left: 167px; position: absolute; top: 1042px" Visible="False"
                   Width="64px">
                   <asp:ListItem>01</asp:ListItem>
                   <asp:ListItem>02</asp:ListItem>
               </asp:DropDownList>
               <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 102; left: 171px; position: absolute; top: 1072px"
                   Visible="False" Width="424px"></asp:TextBox>
               <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
                   Style="z-index: 103; left: 14px; position: absolute; top: 1157px" Text="核定履歷"></asp:Label>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
                   Style="z-index: 104; left: 13px; position: absolute; top: 1179px" Width="780px">
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
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 105; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 106; left: 354px; position: absolute; top: 1144px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 107; left: 445px; position: absolute; top: 1144px" Text="NG2"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 108; left: 537px; position: absolute; top: 1144px" Text="NG1"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 109; left: 629px; position: absolute; top: 1144px" Text="OK"
                   UseSubmitBehavior="false" Width="80px" />
           </div>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
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
           &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/CloseAccountSheet_02.jpg"
               Style="z-index: 111; left: 8px; position: absolute; top: 8px" />
           <asp:Button ID="DAttachfile1" runat="server" CausesValidation="False" Style="z-index: 136;
               left: 1008px; position: absolute; top: 112px" Text="開啟附檔" Width="88px" />
            <asp:TextBox ID="D3" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 137; left: 1264px; position: absolute; top: 216px"
               Width="190px"></asp:TextBox>      &nbsp;
           <input id="BSDate" runat="server" style="z-index: 148; left: 808px; width: 24px;
               position: absolute; top: 240px" type="button" value="..." visible="true" />
           <asp:TextBox ID="DQCNO" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
               Font-Names="Times New Roman" Font-Size="11pt" ForeColor="White" Height="24px"
               Style="z-index: 112; left: 1088px; position: absolute; top: 72px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DDivision3" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 113; left: 152px; position: absolute;
               top: 208px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DDivisionCode3" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 114; left: 240px; position: absolute;
               top: 208px" Width="80px"></asp:TextBox>
           <input id="BAddTrip" runat="server" style="z-index: 149; left: 928px; width: 24px;
               position: absolute; top: 144px" type="button" value="..." />
           <asp:HyperLink ID="LTripNo" runat="server" Style="z-index: 115; left: 960px; position: absolute;
               top: 144px" Target="_blank" Visible="False">Link</asp:HyperLink>
           <asp:HyperLink ID="LQCNo" runat="server" Style="z-index: 116; left: 1000px; position: absolute;
               top: 184px" Target="_blank" Visible="False">[LQCNo]</asp:HyperLink>
           <asp:TextBox ID="DASDate" runat="server" AutoPostBack="True" BackColor="GreenYellow"
               BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 117; left: 712px;
               position: absolute; top: 240px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DAEDate" runat="server" AutoPostBack="True" BackColor="GreenYellow"
               BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 118; left: 840px;
               position: absolute; top: 240px" Width="88px"></asp:TextBox>
           <input id="BEDate" runat="server" style="z-index: 150; left: 936px; width: 24px;
               position: absolute; top: 240px" type="button" value="..." visible="true" />
           <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
               Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Height="24px"
               Style="z-index: 119; left: 712px; position: absolute; top: 112px" Width="144px"></asp:TextBox>
           <asp:TextBox ID="DLocation" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 712px; position: absolute;
               top: 208px; text-align: left" Width="464px"></asp:TextBox>
           <asp:TextBox ID="DTripNo" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 121; left: 712px; position: absolute;
               top: 144px; text-align: left" Width="208px"></asp:TextBox>
           <asp:TextBox ID="DDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" ReadOnly="true" Style="z-index: 122; left: 152px; position: absolute;
               top: 112px" Width="152px"></asp:TextBox>
           <asp:TextBox ID="DDivision1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="true" Style="z-index: 123; left: 152px;
               position: absolute; top: 144px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpName1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="true" Style="z-index: 124; left: 240px;
               position: absolute; top: 144px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DDivisionCode1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="true" Style="z-index: 125; left: 328px;
               position: absolute; top: 144px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpID1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="true" Style="z-index: 126; left: 416px;
               position: absolute; top: 144px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DJobTitle1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="true" Style="z-index: 127; left: 504px;
               position: absolute; top: 144px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="DDivision2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 128; left: 152px; position: absolute;
               top: 176px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpName2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 129; left: 240px; position: absolute;
               top: 176px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DDivisionCode2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 130; left: 328px; position: absolute;
               top: 176px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpID2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 131; left: 416px; position: absolute;
               top: 176px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DJobTitle2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 132; left: 504px; position: absolute;
               top: 176px" Width="64px"></asp:TextBox>
           <asp:Button ID="BAdd" runat="server" Style="z-index: 133; left: 88px;
               position: absolute; top: 288px" Text="登錄" Visible="False" Width="88px" />
           <asp:Button ID="BCheck" runat="server" Style="z-index: 134; left: 976px; position: absolute;
               top: 240px" Text="計算天數" Width="88px" BackColor="#FFC080" />
           <asp:Label ID="Label1" runat="server" Style="z-index: 135; left: 256px; position: absolute;
               top: 248px" Text="~"></asp:Label>
           <asp:TextBox ID="DSDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 136; left: 152px; position: absolute;
               top: 240px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DEDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 137; left: 280px; position: absolute;
               top: 240px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DObject" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 138; left: 712px; position: absolute;
               top: 176px; text-align: left" Width="176px"></asp:TextBox>
           <asp:TextBox ID="DSumAmt" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 139; left: 1024px; position: absolute;
               top: 296px; text-align: right" Width="152px">0</asp:TextBox>
           <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="312px" ScrollBars="Auto"
               Style="z-index: 140; left: 16px; position: absolute; top: 328px" Width="1184px">
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
                       <asp:BoundField DataField="Money" DataFormatString="{0:N2}" HeaderText="金額">
                           <HeaderStyle Height="20px" />
                           <ItemStyle HorizontalAlign="Right" Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Days" HeaderText="數量/次數">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="100px" HorizontalAlign="Right" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Rate" HeaderText="匯率">
                           <HeaderStyle Height="20px" />
                           <ItemStyle HorizontalAlign="Right" Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="SumAmt" HeaderText="小計金額（台幣)" DataFormatString="{0:N0}">
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
               &nbsp; &nbsp;&nbsp; &nbsp;
           </asp:Panel>
           <asp:TextBox ID="D1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 141; left: 1264px; position: absolute; top: 256px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DTNo" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 142; left: 1280px; position: absolute; top: 160px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DUpdate" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 143; left: 1256px; position: absolute; top: 344px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DDays" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 144; left: 1256px; position: absolute; top: 384px"
               Width="142px"></asp:TextBox>
               <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                   ErrorMessage="不可為空白" Width="96px" style="z-index: 151; left: 32px; position: absolute; top: 672px" Height="16px"></asp:RequiredFieldValidator>
           <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
               left: 16px; position: absolute; top: 952px" />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="56px" Style="z-index: 132; left: 64px; position: absolute; top: 960px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
  	 
      </div>
    </form>
</body>
</html>
