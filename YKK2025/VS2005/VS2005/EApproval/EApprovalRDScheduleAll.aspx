<%@ Page Language="vb" AutoEventWireup="false" Inherits="EApprovalRDScheduleAll" aspCompat="True" EnableEventValidation = "false"  CodeFile="EApprovalRDScheduleAll.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>�ݳB�z</title>
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

				answer = confirm("�O�_����[" + MSG + "]�@�~�ܡH �нT�{....");
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
        	<FONT face="�s�ө���"></FONT>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 1104px; POSITION: absolute; TOP: 8px" runat="server"
				ImageUrl="~/Images/msexcel.gif" Height="21px" Width="21px"></asp:imagebutton>
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
               BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
               Style="z-index: 102; left: 8px; position: absolute; top: 8px" Width="1080px">
               <Columns>
                   <asp:BoundField DataField="StsDesc" HeaderText="���A" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="FormName" HeaderText="�e�U��" >
                       <ItemStyle Width="150px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="No" HeaderText="�e�UNo" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="ApplyName" HeaderText="�e�U�H" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="StepNameDesc" HeaderText="�u�{" >
                       <ItemStyle Width="140px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="FlowTypeDesc" HeaderText="���O">
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="AgentName" HeaderText="�N�z/��¾">
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DelaySts" HeaderText="���`/����">
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="Description" HeaderText="�ѦҸ��">
                       <ItemStyle Width="493px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="formsno" HeaderText="formsno" />
                   <asp:BoundField DataField="ApplyID" HeaderText="ApplyID" />
               </Columns>
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:GridView>
  	 
      </div>
    </form>
</body>
</html>
