<%@ Page Language="vb" AutoEventWireup="false" Inherits="CloseAccountSheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="CloseAccountSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
   <title>�M��ӽ�</title>

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
       		
           function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    btn.disabled="disabled";
				    var answer = confirm("�O�_����@�~�ܡH �нT�{....");
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
        	<FONT face="�s�ө���"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <!--************************************************************ 
            ** ����Ы� [BUTTON 2 ��]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 100; left: -500px; position: absolute;
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
               Style="z-index: 101; left: 8px; position: absolute; top: 8px" />
           <asp:Button ID="DAttachfile1" runat="server" CausesValidation="False" Style="z-index: 136;
               left: 1008px; position: absolute; top: 112px" Text="�}�Ҫ���" Width="88px" />
           <asp:HyperLink ID="LPrint" runat="server" Height="1px" NavigateUrl="BoardEdit.aspx"
               Style="z-index: 137; left: 1080px; position: absolute; top: 64px" Target="_blank"
               Width="96px">�C�L���</asp:HyperLink>
           <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 248px; position: absolute;
               top: 248px" Text="~"></asp:Label>
           &nbsp;&nbsp;
           <asp:TextBox ID="DQCNO" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
               Font-Names="Times New Roman" Font-Size="11pt" ForeColor="White" Height="24px"
               Style="z-index: 104; left: 1088px; position: absolute; top: 176px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DDivision3" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 105; left: 152px; position: absolute;
               top: 216px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DDivisionCode3" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 106; left: 240px; position: absolute;
               top: 216px" Width="80px"></asp:TextBox>
           &nbsp;
           <asp:HyperLink ID="LTripNo" runat="server" Style="z-index: 107; left: 960px; position: absolute;
               top: 152px" Target="_blank" Visible="False">Link</asp:HyperLink>
           <asp:HyperLink ID="LQCNo" runat="server" Style="z-index: 108; left: 1000px; position: absolute;
               top: 184px" Target="_blank" Visible="False">[LQCNo]</asp:HyperLink>
           <asp:TextBox ID="DASDate" runat="server" AutoPostBack="True" BackColor="White"
               BorderStyle="Groove" ForeColor="Blue" Height="18px" Style="z-index: 109; left: 712px;
               position: absolute; top: 248px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DAEDate" runat="server" AutoPostBack="True" BackColor="White"
               BorderStyle="Groove" ForeColor="Blue" Height="18px" Style="z-index: 110; left: 840px;
               position: absolute; top: 248px" Width="88px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
               Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Height="24px"
               Style="z-index: 111; left: 704px; position: absolute; top: 112px" Width="176px"></asp:TextBox>
           <asp:TextBox ID="DLocation" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 112; left: 712px; position: absolute;
               top: 216px; text-align: left" Width="456px"></asp:TextBox>
           <asp:TextBox ID="DTripNo" runat="server" BackColor="White" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 113; left: 712px; position: absolute;
               top: 152px; text-align: left" Width="208px"></asp:TextBox>
           <asp:TextBox ID="DDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove" ForeColor="Blue"
               Height="18px" ReadOnly="true" Style="z-index: 114; left: 152px; position: absolute;
               top: 120px" Width="152px"></asp:TextBox>
           <asp:TextBox ID="DDivision1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" ReadOnly="true" Style="z-index: 115; left: 152px;
               position: absolute; top: 152px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpName1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" ReadOnly="true" Style="z-index: 116; left: 240px;
               position: absolute; top: 152px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DDivisionCode1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" ReadOnly="true" Style="z-index: 117; left: 328px;
               position: absolute; top: 152px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpID1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" ReadOnly="true" Style="z-index: 118; left: 416px;
               position: absolute; top: 152px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DJobTitle1" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" ReadOnly="true" Style="z-index: 119; left: 504px;
               position: absolute; top: 152px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="DDivision2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 120; left: 152px; position: absolute;
               top: 184px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpName2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 121; left: 240px; position: absolute;
               top: 184px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DDivisionCode2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 122; left: 328px; position: absolute;
               top: 184px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DEmpID2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 123; left: 416px; position: absolute;
               top: 184px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DJobTitle2" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 124; left: 504px; position: absolute;
               top: 184px" Width="64px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:Label ID="Label1" runat="server" Style="z-index: 125; left: 816px; position: absolute;
               top: 248px" Text="~"></asp:Label>
           <asp:TextBox ID="DSDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 126; left: 152px; position: absolute;
               top: 248px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DEDate" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 127; left: 280px; position: absolute;
               top: 248px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DObject" runat="server" BackColor="#E0E0E0" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 128; left: 712px; position: absolute;
               top: 184px; text-align: left" Width="176px"></asp:TextBox>
           <asp:TextBox ID="DSumAmt" runat="server" BackColor="White" BorderStyle="Groove"
               ForeColor="Blue" Height="18px" Style="z-index: 129; left: 1024px; position: absolute;
               top: 296px; text-align: right" Width="152px">0</asp:TextBox>
           <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="312px" ScrollBars="Auto"
               Style="z-index: 130; left: 16px; position: absolute; top: 328px" Width="1184px">
               &nbsp; &nbsp;&nbsp;
               <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="White"
                   BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" CellPadding="3" DataKeyNames="Unique_ID"
                   Font-Size="9pt" ForeColor="Black" GridLines="Vertical" PageSize="100" Style="z-index: 103;
                   left: 8px; position: absolute; top: 8px" Width="1160px">
                   <Columns>
                       <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Type" HeaderText="���O�γ~">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Pay" HeaderText="��I" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="SDate" HeaderText="���">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Currency" HeaderText="���O">
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Money" DataFormatString="{0:N2}" HeaderText="���B">
                           <ItemStyle HorizontalAlign="Right" Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Days" HeaderText="�ƶq/����">
                           <HeaderStyle Height="20px" />
                           <ItemStyle Width="80px" HorizontalAlign="Right" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Rate" HeaderText="�ײv">
                           <HeaderStyle Height="20px" />
                           <ItemStyle HorizontalAlign="Right" Width="50px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="SumAmt" DataFormatString="{0:N0}" HeaderText="�p�p���B�]�x��)">
                           <ItemStyle HorizontalAlign="Right" Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="ExpenseNo" HeaderText="��ڶO�渹" >
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Remark" HeaderText="�Ƶ�">
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
               Height="24px" Style="z-index: 131; left: 1264px; position: absolute; top: 256px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DTNo" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 132; left: 1280px; position: absolute; top: 160px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DUpdate" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 133; left: 1256px; position: absolute; top: 344px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DDays" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 134; left: 1256px; position: absolute; top: 384px"
               Width="142px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp;&nbsp;
           <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
               Style="z-index: 135; left: 24px; position: absolute; top: 664px" Text="�֩w�i��"></asp:Label>
           <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
               BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
               Style="z-index: 136; left: 24px; position: absolute; top: 680px" Width="780px">
               <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
               <Columns>
                   <asp:BoundField DataField="StepNameDesc" HeaderText="�u�{">
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DecideName" HeaderText="���" />
                   <asp:BoundField DataField="AgentName" HeaderText="�N�z/��¾" />
                   <asp:BoundField DataField="FlowTypeDesc" HeaderText="���O" />
                   <asp:BoundField DataField="StsDesc" HeaderText="�B�z���G" />
                   <asp:BoundField DataField="DecideDescA" HeaderText="����">
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundField>
                   <asp:BoundField DataField="Description" HeaderText="�֩w�ɶ�">
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundField>
               </Columns>
               <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                   VerticalAlign="Middle" />
           </asp:GridView>
  	 
      </div>
    </form>
</body>
</html>
