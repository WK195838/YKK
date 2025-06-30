<%@ Page Language="vb" AutoEventWireup="false" Inherits="SBDSurfaceSheet_02" aspCompat="True" CodeFile="SBDSurfaceSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>S & B開發委託書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/SBD-SufaceSheet01.jpg" style="z-index: 99; left: 10px; position: absolute; top: 15px" />
           <asp:TextBox ID="DAttachSample" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 101; left: 632px; position: absolute;
               top: 230px" Width="168px" ReadOnly="True"></asp:TextBox>
           <asp:Image ID="LCustSampleFile" runat="server" BorderStyle="Groove" Height="256px"
               ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
               Style="z-index: 102; left: 24px; position: absolute; top: 128px" Width="200px" />
           <asp:Image ID="LFinalSampleFile" runat="server" BorderStyle="Groove" Height="192px"
               ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
               Style="z-index: 103; left: 24px; position: absolute; top: 432px" Width="200px" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LForcastFile" runat="server" Style="z-index: 104; left: 528px; position: absolute;
               top: 720px" Target="_blank" Width="168px">報價單</asp:HyperLink>
           <asp:HyperLink ID="LQAFile" runat="server" Style="z-index: 105; left: 344px; position: absolute;
               top: 608px" Target="_blank" Width="368px">品測報告書</asp:HyperLink>
           &nbsp; &nbsp;
           <asp:HyperLink ID="LContactFile" runat="server" Style="z-index: 106; left: 528px; position: absolute;
               top: 752px" Target="_blank" Width="168px">切結書</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LManufFlowFile" runat="server" Style="z-index: 107; left: 136px;
               position: absolute; top: 720px" Target="_blank" Width="176px">製造流程表</asp:HyperLink>
           <asp:HyperLink ID="LOPManualFile" runat="server" Style="z-index: 108; left: 136px;
               position: absolute; top: 752px" Target="_blank" Width="184px">作業標準書</asp:HyperLink>
           &nbsp;
           <asp:HyperLink ID="LQCReqFile" runat="server" Style="z-index: 109; left: 344px; position: absolute;
               top: 568px" Target="_blank" Width="360px">品質依賴書</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 110; left: 336px; position: absolute; top: 160px"
               Width="176px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <input id="DQCDate" runat="server" style="z-index: 131; left: 112px; width: 96px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 672px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" readonly="readOnly" />
           &nbsp;
           <input id="DBFinalDate" runat="server" style="z-index: 132; left: 336px; width: 120px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 472px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" readonly="readOnly" />
           &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:TextBox ID="DSellVendor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 111; left: 632px; position: absolute;
            top: 164px" Width="168px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 112; left: 336px; position: absolute;
            top: 128px" Width="176px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DReqQty" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 113; left: 632px; position: absolute; top: 198px"
            Width="168px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 114; left: 336px; position: absolute; top: 230px"
               Width="176px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DORNO" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 115; left: 336px; position: absolute; top: 264px"
               Width="176px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DColor" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 116; left: 336px; position: absolute; top: 328px"
               Width="176px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DSampleQty" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 117; left: 632px; position: absolute;
               top: 328px" Width="176px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           &nbsp;&nbsp;<input id="DOrderTime" runat="server" style="z-index: 130; left: 632px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 264px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" readonly="readOnly" /><input id="DReqDelDate" runat="server" style="z-index: 129; left: 336px; width: 112px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 200px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" readonly="readOnly" />
           &nbsp;&nbsp;
        <asp:TextBox ID="DDevReason" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 118; left: 336px; position: absolute;
            top: 296px" Width="464px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DCap" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 119; left: 624px; position: absolute;
            top: 436px" Width="120px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DAllowSample" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 336px; position: absolute;
               top: 436px" Width="120px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DQCResult" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 121; left: 240px; position: absolute;
               top: 672px" Width="80px" ReadOnly="True"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DSchedule" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 122; left: 624px; position: absolute;
               top: 472px" Width="120px" ReadOnly="True"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DEnglishName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 123; left: 336px; position: absolute;
               top: 360px" Width="464px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DSupplier" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 124; left: 336px; position: absolute;
               top: 400px" Width="464px" ReadOnly="True"></asp:TextBox>
           &nbsp;&nbsp;
        <asp:TextBox ID="DQCRemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 125; left: 336px; position: absolute;
            top: 672px" Width="464px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAppper" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 126; left: 632px; position: absolute;
            top: 128px" Width="168px" ReadOnly="True"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 127; left: 632px; position: absolute;
            top: 96px" Width="168px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 128; left: 336px; position: absolute;
            top: 96px" Width="176px" ReadOnly="True"></asp:TextBox>
           &nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
        <br />
           &nbsp; &nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        &nbsp;<br />
        <br />
        <br />
        <br />
           &nbsp;&nbsp;<br />
           &nbsp; &nbsp;
        &nbsp;&nbsp;<br />
           &nbsp;
        </div>
    </form>
</body>
</html>
