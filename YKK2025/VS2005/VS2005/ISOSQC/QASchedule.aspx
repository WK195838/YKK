<%@ Page Language="vb" AutoEventWireup="false" Inherits="QASchedule" aspCompat="True" EnableEventValidation = "false"  CodeFile="QASchedule.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>品質分析依賴</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			
			   function RunExcelCHECK()
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\ISOSQC\\MIS\\ISOSQC_QACHECK.exe';
    try
    {
        var shellActiveXObject = new ActiveXObject("WScript.Shell");

        if ( !shellActiveXObject )
        {
            alert('Could not get reference to WScript.Shell');
        }

        shellActiveXObject.Run(executableFullPath, 1, false);
        shellActiveXObject = null;
    }
    catch (errorObject)
    {
        alert('Error:\n' + errorObject.message);
    }            
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/NewCommission_03.jpg" Style="z-index: 100;
               left: 16px; position: absolute; top: 32px" />
           <asp:HyperLink ID="LEDX4" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 101; left: 168px; position: absolute; top: 360px" Target="_self"
               Width="88px">當年以前</asp:HyperLink>
           <asp:HyperLink ID="LEDX3" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 102; left: 168px; position: absolute; top: 336px" Target="_self"
               Width="56px">當年</asp:HyperLink>
           <asp:HyperLink ID="LEDX2" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 103; left: 24px; position: absolute; top: 360px" Target="_self"
               Width="88px">當年以前</asp:HyperLink>
           <asp:HyperLink ID="LEDX1" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 104; left: 24px; position: absolute; top: 336px" Target="_self"
               Width="56px">當年</asp:HyperLink>
           <asp:HyperLink ID="LEDX6" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 105; left: 304px; position: absolute; top: 360px" Target="_self"
               Width="88px">當年以前</asp:HyperLink>
           <asp:HyperLink ID="LEDX5" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
               Style="z-index: 106; left: 304px; position: absolute; top: 336px" Target="_self"
               Width="56px">當年</asp:HyperLink>
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
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp; 
           <asp:Label ID="Label3" runat="server" Font-Size="Large" Style="z-index: 107; left: 464px;
               position: absolute; top: 416px" Text="品質分析依賴" Width="296px" ForeColor="Blue"></asp:Label>
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Label ID="Label1" runat="server" Style="z-index: 107; left: 456px; position: absolute;
               top: 8px" Text="NOTIC LIST" Font-Size="Large" Width="296px" Font-Bold="True" ForeColor="Red"></asp:Label>
           <asp:Label ID="Label2" runat="server" Style="z-index: 108; left: 464px; position: absolute;
               top: 656px" Text="資料修改依賴" Font-Size="Large" Width="296px" ForeColor="Blue"></asp:Label>
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LFun02" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 109; left: 32px; position: absolute;
               top: 136px" Target="_self" Width="224px">品質分析依賴調閱資料</asp:HyperLink>
           <asp:HyperLink ID="LFun01" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 110; left: 32px; position: absolute;
               top: 104px" Target="_self" Width="224px">品質分析依賴新委託</asp:HyperLink><asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="10pt" ShowFooter="True"
               Style="z-index: 111; left: 464px; position: absolute; top: 448px" Width="688px" HorizontalAlign="Center" Height="144px">
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
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView>
               <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="10pt" ShowFooter="True"
               Style="z-index: 112; left: 464px; position: absolute; top: 688px" Width="688px" HorizontalAlign="Center">
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
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               </asp:GridView>
           &nbsp; &nbsp;
           <asp:HyperLink ID="LFun03" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 113; left: 32px; position: absolute;
               top: 232px" Target="_self" Width="224px" Visible="False">品質分析依賴EDX判定調閱</asp:HyperLink>
           <asp:HyperLink ID="LFun04" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 114; left: 32px; position: absolute;
               top: 200px" Target="_self" Width="224px">資料修改調閱資料</asp:HyperLink>
           <asp:HyperLink ID="LFun05" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 115; left: 32px; position: absolute;
               top: 168px" Target="_self" Width="224px">資料修改新委託</asp:HyperLink>
           &nbsp; &nbsp;&nbsp; &nbsp;
           <asp:Panel ID="Panel1" runat="server" Height="352px" Style="z-index: 117; left: 448px;
               position: absolute; top: 32px" Width="1384px" ScrollBars="Vertical">
               <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderWidth="1px" CellPadding="4" Font-Size="10pt" ShowFooter="True"
               Style="z-index: 111; left: 8px; position: absolute; top: 8px" Width="1336px" HorizontalAlign="Center">
                   <Columns>
                       <asp:HyperLinkField DataNavigateUrlFields="ViewURL" DataTextField="Field1" Text="NO" >
                           <ItemStyle Width="80px" />
                       </asp:HyperLinkField>
                       <asp:BoundField DataField="Field2" HeaderText="Field2" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field3" HeaderText="Field3" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field4" HeaderText="Field4" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field5" HeaderText="Field5" >
                           <ItemStyle Width="120px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field6" HeaderText="Field6" >
                           <ItemStyle Width="100px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field7" HeaderText="Field7" >
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field8" HeaderText="Field8" >
                           <ItemStyle Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Field9" HeaderText="Field9" >
                           <ItemStyle Width="120px" />
                       </asp:BoundField>
                   </Columns>
                   <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                   <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
                   <EmptyDataTemplate>
                       暫無資料
                   </EmptyDataTemplate>
               </asp:GridView>
           </asp:Panel>
       </div>
    </form>
</body>
</html>
