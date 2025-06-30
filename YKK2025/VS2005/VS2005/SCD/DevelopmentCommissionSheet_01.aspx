<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentCommissionSheet_01.aspx.vb" Inherits="DevelopmentCommissionSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>開發委託書</title>
	    <script language="javascript" src="SubProgram.js"></script>

		<script language="javascript" type="text/javascript">
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
    <form id="form1" runat="server">
    <div>
        <!--************************************************************ 
            ** Panel Start                                            **     
            ************************************************************ -->
        <asp:Panel ID="Panel1" runat="server" Style="left: 0px; position: relative; top: 0px" Width="694px">
          <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/Images/DevelopmentCommission_Blank.jpg" />
          <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="~/Images/ManufactureCommission_Button.jpg" style="left: -6px; position: relative; top: 0px" />
          <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="~/Images/DevelopmentSample_Button.jpg" style="left: -14px; position: relative; top: 0px" />
          <asp:ImageButton ID="ImageButton4" runat="server" ImageUrl="~/Images/DevelopmentGentani_Button.jpg" style="left: -22px; position: relative; top: 0px" />
          <asp:ImageButton ID="ImageButton5" runat="server" ImageUrl="~/Images/SignHistory_Button.jpg" style="left: -30px; position: relative; top: 0px" />
        </asp:Panel>
        <!-- Panel End -->
        <!-- -->
        <!--************************************************************ 
            ** MultiView Start                                           
            ************************************************************ -->
            <asp:MultiView ID="MultiView1" runat="server" ActiveViewIndex="0" >
            <!-- -->
            <!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 開發委託書(View1) Start                 ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->
              <asp:View ID="View1" runat="server">
<table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">

                     <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DevelopmentCommission_01.png"
                     Style="z-index: 99; left: 2px; position: absolute; top: -1px" />

        <asp:FileUpload ID="DMAPFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 124px; position: absolute; top: 1197px; background-color: #c0ffff"
            Width="644px" />
        <asp:FileUpload ID="DFORTYPEFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 124px; position: absolute; top: 1257px; background-color: #c0ffff"
            Width="644px" />
        <asp:FileUpload ID="DMAPREFFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 124px; position: absolute; top: 236px; background-color: #c0ffff"
            Width="644px" />
        <asp:HyperLink ID="LFORTYPEFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 261; left: 126px; position: absolute; top: 1259px" Target="_blank"
            Width="104px">適用型別檔</asp:HyperLink>
        &nbsp;
        <asp:HyperLink ID="LMAPREFFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 261; left: 125px; position: absolute; top: 239px" Target="_blank"
            Width="64px">草圖</asp:HyperLink>
        &nbsp;
        <asp:DropDownList ID="DLEVEL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 1170px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DMAKEMAP" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 1144px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DMAPNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 124px;
            position: absolute; top: 1144px; text-align: left" Width="264px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOTCON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 124px;
            position: absolute; top: 1084px; text-align: left" Width="643px" ReadOnly="True"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DLYLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 1032px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DLYYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 1032px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DLYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 1032px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DLYCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 1032px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DHMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 1006px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DHMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 1006px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DHMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 1006px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DHMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 1006px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DGMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 980px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DGMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 980px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DGMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 980px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DGMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 980px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DFMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 954px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DFMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 954px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DFMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 954px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DFMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 954px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DEMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 928px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DEMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 928px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DEMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 928px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DEMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 928px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DDMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 902px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DDMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 902px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DDMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 902px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DDMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 902px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DCMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 876px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DCMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 876px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DCMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 876px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DCMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 876px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DBMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 850px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DBMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 850px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DBMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 850px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DBMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 850px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DAMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 824px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 824px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 824px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DAMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 824px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DXMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 375px;
            position: absolute; top: 798px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DXMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 627px;
            position: absolute; top: 798px; text-align: left" Width="140px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DXMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 550px;
            position: absolute; top: 798px; text-align: left" Width="70px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DXMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 479px; position: absolute; top: 798px"
            Width="70px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHRLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 565px;
            position: absolute; top: 746px; text-align: left" Width="201px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHLLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 335px;
            position: absolute; top: 746px; text-align: left" Width="221px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 124px;
            position: absolute; top: 746px; text-align: left" Width="201px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHRLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 666px;
            position: absolute; top: 720px; text-align: left" Width="100px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTHRLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 565px; position: absolute; top: 720px"
            Width="100px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 446px;
            position: absolute; top: 720px; text-align: left" Width="110px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTHLLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 334px; position: absolute; top: 720px"
            Width="110px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 225px;
            position: absolute; top: 720px; text-align: left" Width="100px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTHLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 720px"
            Width="100px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHRUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 565px;
            position: absolute; top: 694px; text-align: left" Width="201px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHLUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 335px;
            position: absolute; top: 694px; text-align: left" Width="221px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 124px;
            position: absolute; top: 694px; text-align: left" Width="201px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHRUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 666px;
            position: absolute; top: 668px; text-align: left" Width="100px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTHRUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 565px; position: absolute; top: 668px"
            Width="100px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 446px;
            position: absolute; top: 668px; text-align: left" Width="110px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTHLUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 334px; position: absolute; top: 668px"
            Width="110px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 225px;
            position: absolute; top: 668px; text-align: left" Width="100px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTHUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 668px"
            Width="100px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTARYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 565px;
            position: absolute; top: 642px; text-align: left" Width="201px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTALYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 335px;
            position: absolute; top: 642px; text-align: left" Width="221px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTAYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 124px;
            position: absolute; top: 642px; text-align: left" Width="201px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTARCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 666px;
            position: absolute; top: 616px; text-align: left" Width="100px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTARCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 565px; position: absolute; top: 616px"
            Width="100px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTALCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 446px;
            position: absolute; top: 616px; text-align: left" Width="110px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTALCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 334px; position: absolute; top: 616px"
            Width="110px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTACOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 225px;
            position: absolute; top: 616px; text-align: left" Width="100px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTACOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 616px"
            Width="100px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DCCOL" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 634px;
            position: absolute; top: 564px; text-align: left" Width="131px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DCCOLSEL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 564px"
            Width="130px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DECOL" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 256px;
            position: absolute; top: 564px; text-align: left" Width="131px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DTATYPE" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 538px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DSIZENO" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 512px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DITEM" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 512px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DNEEDMAP" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 210px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTAWIDTH" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 502px;
            position: absolute; top: 538px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DSPSPEC" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 124px;
            position: absolute; top: 452px; text-align: left" Width="641px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DLOSTK" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 565px;
            position: absolute; top: 426px; text-align: left" Width="200px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DUPSTK" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 565px;
            position: absolute; top: 400px; text-align: left" Width="200px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DLOFIN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 334px;
            position: absolute; top: 426px; text-align: left" Width="222px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DUPFIN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 334px;
            position: absolute; top: 400px; text-align: left" Width="222px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DLOSLI" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 124px;
            position: absolute; top: 426px; text-align: left" Width="200px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DUPSLI" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 124px;
            position: absolute; top: 400px; text-align: left" Width="200px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DEAQTY" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 607px; position: absolute;
            top: 348px; text-align: left" Width="74px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DEAQTYUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 684px; position: absolute; top: 348px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DPQTY" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 607px; position: absolute;
            top: 322px; text-align: left" Width="74px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DPQTYUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 684px; position: absolute; top: 322px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DEALEN" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 439px; position: absolute;
            top: 348px; text-align: left" Width="74px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DEALENUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 516px; position: absolute; top: 348px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DPRO" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 123px; position: absolute; top: 296px"
            Width="204px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DPLEN" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 439px; position: absolute;
            top: 322px; text-align: left" Width="74px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOPPART" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 322px; text-align: left" Width="200px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DECOLSEL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 564px"
            Width="130px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:HyperLink ID="LMAPFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 261; left: 126px; position: absolute; top: 1199px" Target="_blank"
            Width="64px">圖檔</asp:HyperLink>
        &nbsp;
        <asp:DropDownList ID="DUSAGE" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 184px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DAPPBUYER" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 132px"
            Width="266px">
            <asp:ListItem Value="徐">徐</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DORNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 210px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DCUSTITEM" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 184px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DEXPDEL" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="10" Style="z-index: 126; left: 502px; position: absolute;
            top: 158px; text-align: left" Width="240px" ReadOnly="True"></asp:TextBox>
        <asp:Button ID="BEXPDEL" runat="server" Height="20px" Style="z-index: 111; left: 747px;
            position: absolute; top: 160px" Text="....." Width="20px" />
        <asp:TextBox ID="DESYQTY" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 158px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DSellVendor" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="10" Style="z-index: 126; left: 502px;
            position: absolute; top: 132px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DPLENUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 516px; position: absolute; top: 322px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DAPPPER" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="10" Style="z-index: 126; left: 502px; position: absolute;
            top: 106px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAPPDEPT" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 124px;
            position: absolute; top: 106px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAPPDATE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="10" Style="z-index: 126; left: 502px; position: absolute;
            top: 80px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DRNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 80px; text-align: left" Width="263px" ReadOnly="True"></asp:TextBox>

</table>
              </asp:View>
            <!-- View-1 End -->
            <!-- -->
            <!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 製造委託書(View2) Start                 ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->
              <asp:View ID="View2" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">

                     <asp:Image ID="Image2" runat="server" ImageUrl="~/Images/ManufactureCommission_01.PNG"
                     Style="z-index: 99; left: 2px; position: absolute; top: -1px" />
        <asp:TextBox ID="DOP8REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 1292px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP8DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1318px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP8DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1292px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP8AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 1266px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP8ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 1266px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP8BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 1240px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP8CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 1240px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP8BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 1240px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP8" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 1240px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP8PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 1240px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 1188px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP7DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1214px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP7DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1188px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP7AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 1162px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 1162px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 1136px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 1136px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 1136px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 1136px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP7PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 1136px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 1084px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP6DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1110px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP6DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1084px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP6AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 1058px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 1058px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 1032px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 1032px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 1032px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 1032px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP6PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 1032px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 980px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP5DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1006px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP5DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 980px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP5AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 954px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 954px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 928px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 928px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 928px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 928px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP5PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 928px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 876px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP4DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 902px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP4DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 876px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP4AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 850px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 850px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 824px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 824px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 824px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 824px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP4PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 824px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 772px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP3DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 798px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP3DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 772px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP3AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 746px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 746px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 720px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 720px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 720px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 720px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP3PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 720px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 668px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP2DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 694px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP2DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 668px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP2AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 642px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 642px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 616px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 616px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 616px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 616px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP2PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 616px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 564px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:DropDownList ID="DOP1DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 590px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP1DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 564px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTASPEC" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="99px" MaxLength="14" Style="z-index: 126; left: 333px;
            position: absolute; top: 243px; text-align: left" TextMode="MultiLine" Width="433px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 538px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 538px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 354px;
            position: absolute; top: 512px; text-align: left" Width="56px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 564px;
            position: absolute; top: 374px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 333px;
            position: absolute; top: 348px; text-align: left" Width="119px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px" MaxLength="14" Style="z-index: 126; left: 416px;
            position: absolute; top: 512px; text-align: left" TextMode="MultiLine" Width="351px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DTHSPEC" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="47px" MaxLength="14" Style="z-index: 126; left: 123px; position: absolute;
            top: 373px; text-align: left" TextMode="MultiLine" Width="329px" ReadOnly="True"></asp:TextBox>
        <input id="DHINTFILE" runat="server" name="File1" style="z-index: 260; left: 20px; width: 201px;
            position: absolute; top: 323px; height: 20px;" type="file" />
        <asp:Image ID="LHINTFILE" runat="server" BorderStyle="Groove" Height="145px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
            Style="z-index: 241; left: 20px; position: absolute; top: 194px" Width="200px" />
        <asp:TextBox ID="DDEVTITLE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 80px; text-align: left" Width="642px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DMANUFTYPE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 460px; text-align: left" Width="642px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DDEVNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 106px; text-align: left" Width="264px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DCODENO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 502px; position: absolute;
            top: 106px; text-align: left" Width="264px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DDEVPER1" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 502px; position: absolute;
            top: 132px; text-align: left" Width="264px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DISSDATE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 124px; position: absolute;
            top: 132px; text-align: left" Width="264px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 333px; position: absolute;
            top: 192px; text-align: left" Width="160px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 207px;
            position: absolute; top: 512px; text-align: left" Width="143px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 564px;
            position: absolute; top: 400px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox5" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 669px;
            position: absolute; top: 400px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox6" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="14" Style="z-index: 126; left: 669px;
            position: absolute; top: 374px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox7" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 123px; position: absolute;
            top: 348px; text-align: left" Width="98px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 17px;
            position: absolute; top: 512px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DOP1PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px" MaxLength="14" Style="z-index: 126; left: 101px;
            position: absolute; top: 512px; text-align: left" Width="79px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox8" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 333px; position: absolute;
            top: 218px; text-align: left" Width="433px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox9" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="14" Style="z-index: 126; left: 606px; position: absolute;
            top: 192px; text-align: left" Width="160px" ReadOnly="True"></asp:TextBox>

                </table>
              </asp:View>
            <!-- View-2 End -->
            <!-- -->
            <!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 開發見本(View3) Start                     ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->
              <asp:View ID="View3" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">
                    
                    <asp:Image ID="Image4" runat="server" ImageUrl="~/Images/DevelopmentSample_01.png" Style="z-index: 99;
                    left: 5px; position: absolute; top: 2px" />

        <asp:FileUpload ID="D3QCFILE2" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 525px; position: absolute; top: 722px; background-color: #c0ffff"
            Width="267px" />
        <asp:FileUpload ID="D3QCFILE4" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 525px; position: absolute; top: 752px; background-color: #c0ffff"
            Width="267px" />
        <asp:FileUpload ID="D3SAMPLEFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 20px; position: absolute; top: 198px; background-color: #c0ffff"
            Width="771px" />
        <asp:FileUpload ID="D3QCFILE1" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 146px; position: absolute; top: 722px; background-color: #c0ffff"
            Width="248px" />
        <asp:FileUpload ID="D3QCFILE3" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 146px; position: absolute; top: 752px; background-color: #c0ffff"
            Width="248px" />
        <asp:FileUpload ID="D3QCFILE5" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 146px; position: absolute; top: 780px; background-color: #c0ffff"
            Width="248px" />
        &nbsp;
				<asp:hyperlink id="LQCFILE5" style="Z-INDEX: 178; POSITION: absolute; TOP: 781px; LEFT: 148px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">其它</asp:hyperlink><asp:label ID="D31Other" style="Z-INDEX: 173; POSITION: absolute; TOP: 966px; LEFT: 170px" runat="server"
					Font-Size="12px" Width="116px"> </asp:label><asp:label ID="D32Other" style="Z-INDEX: 172; POSITION: absolute; TOP: 990px; LEFT: 170px" runat="server"
					Font-Size="12px" Width="116px"> </asp:label><asp:button id="BCREATE" style="Z-INDEX: 170; POSITION: absolute; TOP: 97px; LEFT: 543px" runat="server"
					Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:dropdownlist ID="D3WF7Name" style="Z-INDEX: 169; POSITION: absolute; TOP: 1080px; LEFT: 674px"
					runat="server" Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF6Name" style="Z-INDEX: 168; POSITION: absolute; TOP: 1080px; LEFT: 569px"
					runat="server" Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF5Name" style="Z-INDEX: 167; POSITION: absolute; TOP: 1080px; LEFT: 464px"
					runat="server" Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF4Name" style="Z-INDEX: 166; POSITION: absolute; TOP: 1080px; LEFT: 359px"
					runat="server" Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF3Name" style="Z-INDEX: 165; POSITION: absolute; TOP: 1080px; LEFT: 254px"
					runat="server" Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQCFILE1" style="Z-INDEX: 160; POSITION: absolute; TOP: 724px; LEFT: 147px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">吋法檔案</asp:hyperlink><asp:dropdownlist ID="D3WF7" style="Z-INDEX: 159; POSITION: absolute; TOP: 1107px; LEFT: 674px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF6" style="Z-INDEX: 158; POSITION: absolute; TOP: 1107px; LEFT: 569px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF5" style="Z-INDEX: 157; POSITION: absolute; TOP: 1107px; LEFT: 464px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF4" style="Z-INDEX: 156; POSITION: absolute; TOP: 1107px; LEFT: 359px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF3" style="Z-INDEX: 155; POSITION: absolute; TOP: 1107px; LEFT: 254px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF2" style="Z-INDEX: 154; POSITION: absolute; TOP: 1107px; LEFT: 149px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist ID="D3WF1" style="Z-INDEX: 109; POSITION: absolute; TOP: 1107px; LEFT: 44px" runat="server"
					Font-Size="9pt" Width="74px" Height="22px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:textbox ID="D3CITEM" style="Z-INDEX: 147; POSITION: absolute; TOP: 1011px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3O2ITEM" style="Z-INDEX: 145; POSITION: absolute; TOP: 987px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3O1ITEM" style="Z-INDEX: 146; POSITION: absolute; TOP: 963px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CDITEM" style="Z-INDEX: 144; POSITION: absolute; TOP: 939px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CSITEM" style="Z-INDEX: 143; POSITION: absolute; TOP: 915px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TDRITEM" style="Z-INDEX: 142; POSITION: absolute; TOP: 867px; LEFT: 525px"
					runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TSRITEM" style="Z-INDEX: 141; POSITION: absolute; TOP: 843px; LEFT: 525px"
					runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TNRITEM" style="Z-INDEX: 140; POSITION: absolute; TOP: 819px; LEFT: 525px"
					runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CNITEM" style="Z-INDEX: 139; POSITION: absolute; TOP: 891px; LEFT: 295px" runat="server"
					Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TDLITEM" style="Z-INDEX: 138; POSITION: absolute; TOP: 867px; LEFT: 295px"
					runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TSLITEM" style="Z-INDEX: 137; POSITION: absolute; TOP: 843px; LEFT: 295px"
					runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TNLITEM" style="Z-INDEX: 136; POSITION: absolute; TOP: 819px; LEFT: 295px"
					runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox>
				<asp:hyperlink id="LQCFILE2" style="Z-INDEX: 161; POSITION: absolute; TOP: 724px; LEFT: 524px"
					runat="server" Font-Size="12pt" Height="8px" Target="_blank" Width="121px">強度檔案</asp:hyperlink><asp:hyperlink id="LQCFILE3" style="Z-INDEX: 163; POSITION: absolute; TOP: 754px; LEFT: 147px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">往覆測試檔案</asp:hyperlink><asp:hyperlink id="LQCFILE4" style="Z-INDEX: 162; POSITION: absolute; TOP: 751px; LEFT: 524px"
					runat="server" Font-Size="12pt" Height="8px" Target="_blank" Width="188px">仕樣書.組織圖檔案</asp:hyperlink>
				<asp:textbox ID="D3OTHER" style="Z-INDEX: 130; POSITION: absolute; TOP: 634px; LEFT: 147px" runat="server"
					Width="640px" Height="76px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240"
					TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3THCOL" style="Z-INDEX: 129; POSITION: absolute; TOP: 578px; LEFT: 147px" runat="server"
					Width="640px" Height="45px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3CCOL" style="Z-INDEX: 128; POSITION: absolute; TOP: 547px; LEFT: 147px" runat="server"
					Width="640px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3ECOL" style="Z-INDEX: 127; POSITION: absolute; TOP: 518px; LEFT: 147px" runat="server"
					Width="640px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TACOL" style="Z-INDEX: 126; POSITION: absolute; TOP: 402px; LEFT: 147px" runat="server"
					Width="640px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3DEVPRD" style="Z-INDEX: 125; POSITION: absolute; TOP: 372px; LEFT: 609px" runat="server"
					Width="178px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3DEVNO" style="Z-INDEX: 124; POSITION: absolute; TOP: 372px; LEFT: 378px" runat="server"
					Width="115px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TAWIDTH" style="Z-INDEX: 123; POSITION: absolute; TOP: 372px; LEFT: 147px"
					runat="server" Width="85px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox>
				<asp:image id="LSAMPLEFILE" style="Z-INDEX: 121; POSITION: absolute; TOP: 198px; LEFT: 21px"
					runat="server" Width="770px" Height="168px" BorderStyle="Groove" ImageUrl="F:\DMF04006-DS2W.jpg"></asp:image><asp:textbox ID="D3CODENO" style="Z-INDEX: 120; POSITION: absolute; TOP: 144px; LEFT: 566px" runat="server"
					Width="220px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3ITEM" style="Z-INDEX: 119; POSITION: absolute; TOP: 144px; LEFT: 356px" runat="server"
					Width="200px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3SIZENO" style="Z-INDEX: 118; POSITION: absolute; TOP: 144px; LEFT: 146px" runat="server"
					Width="200px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3TALINE" style="Z-INDEX: 103; POSITION: absolute; TOP: 460px; LEFT: 147px" runat="server"
					Width="640px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3DATE" style="Z-INDEX: 164; POSITION: absolute; TOP: 91px; LEFT: 671px" runat="server"
					Width="115px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox><asp:textbox ID="D3APPBUYER" style="Z-INDEX: 104; POSITION: absolute; TOP: 91px; LEFT: 147px"
					runat="server" Width="390px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="True"> </asp:textbox>&nbsp;

                </table>
                    
              </asp:View>
            <!-- View-3 End -->
            <!-- -->
            <!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 原單位(View4) Start                     ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->
              <asp:View ID="View4" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">
                  
                    <asp:Image ID="Image3" runat="server" ImageUrl="~/Images/DevelopmentGentani_01.png" Style="z-index: 99;
                    left: 5px; position: absolute; top: 2px" />
                    
     <asp:Label ID="D4RNO" runat="server" Style="z-index: 100; left: 21px; position: absolute;
            top: 92px" Width="70px"></asp:Label>
        <asp:Label ID="D4DEVNO" runat="server" Style="z-index: 101; left: 95px; position: absolute;
            top: 92px" Width="106px"></asp:Label>
        <asp:Label ID="D4SIZENO" runat="server" Style="z-index: 102; left: 206px; position: absolute;
            top: 92px" Width="9px"></asp:Label>
        <asp:Label ID="D4ITEM" runat="server" Style="z-index: 103; left: 280px; position: absolute;
            top: 92px"></asp:Label>
        <asp:Label ID="D4CODENO" runat="server" Style="z-index: 104; left: 353px; position: absolute;
            top: 92px"></asp:Label>
        <asp:Label ID="D4MANUFTYPE" runat="server" Style="z-index: 105; left: 428px; position: absolute;
            top: 92px"></asp:Label>
        <asp:Label ID="D4TALCOL" runat="server" Style="z-index: 106; left: 94px; position: absolute;
            top: 180px"></asp:Label>
        <asp:Label ID="D4XMLCOL" runat="server" Style="z-index: 107; left: 94px; position: absolute;
            top: 202px"></asp:Label>
        <asp:Label ID="D4TALCOLNO" runat="server" Style="z-index: 108; left: 170px; position: absolute;
            top: 180px"></asp:Label>
        <asp:Label ID="D4TALITEM" runat="server" Style="z-index: 109; left: 245px; position: absolute;
            top: 180px"></asp:Label>
        <asp:Label ID="D4TALLEN" runat="server" Style="z-index: 110; left: 322px; position: absolute;
            top: 180px"></asp:Label>
        <asp:Label ID="D4TARCOL" runat="server" Style="z-index: 111; left: 433px; position: absolute;
            top: 180px" Text=""></asp:Label>
        <asp:Label ID="D4TARCOLNO" runat="server" Style="z-index: 112; left: 506px; position: absolute;
            top: 180px" Text=""></asp:Label>
        <asp:Label ID="D4TARITEM" runat="server" Style="z-index: 113; left: 579px; position: absolute;
            top: 180px"></asp:Label>
        <asp:Label ID="D4TCRLEN" runat="server" Style="z-index: 114; left: 658px; position: absolute;
            top: 422px" Text=""></asp:Label>
        <asp:Label ID="D4TARLEN" runat="server" Style="z-index: 115; left: 658px; position: absolute;
            top: 180px" Text=""></asp:Label>
        <asp:Label ID="D4YAMRSUM" runat="server" Style="z-index: 116; left: 653px; position: absolute;
            top: 444px" Text=""></asp:Label>
        <asp:Label ID="D4PICOL" runat="server" Style="z-index: 117; left: 94px; position: absolute;
            top: 510px" Text=""></asp:Label>
        <asp:Label ID="D4CO1ITEM" runat="server" Style="z-index: 118; left: 245px; position: absolute;
            top: 532px" Text=""></asp:Label>
        <asp:Label ID="D4TSLITEM" runat="server" Style="z-index: 119; left: 539px; position: absolute;
            top: 511px" Text=""></asp:Label>
        <asp:Label ID="D4VMCORUNIT" runat="server" Style="z-index: 120; left: 649px; position: absolute;
            top: 620px" Text=""></asp:Label>
        <asp:Label ID="D4TDRITEM" runat="server" Style="z-index: 121; left: 539px; position: absolute;
            top: 620px" Text=""></asp:Label>
        <asp:Label ID="D4VMCOLUNIT" runat="server" Style="z-index: 122; left: 649px; position: absolute;
            top: 598px" Text=""></asp:Label>
        <asp:Label ID="D4TDLITEM" runat="server" Style="z-index: 123; left: 539px; position: absolute;
            top: 598px" Text=""></asp:Label>
        <asp:Label ID="D4VMSETRUNIT" runat="server" Style="z-index: 124; left: 649px; position: absolute;
            top: 533px" Text=""></asp:Label>
        <asp:Label ID="D4VMSETLUNIT" runat="server" Style="z-index: 125; left: 649px; position: absolute;
            top: 511px" Text=""></asp:Label>
        <asp:Label ID="D4TSRITEM" runat="server" Style="z-index: 126; left: 539px; position: absolute;
            top: 533px" Text=""></asp:Label>
        <asp:Label ID="D4CO5LEN" runat="server" Style="z-index: 127; left: 322px; position: absolute;
            top: 620px" Text=""></asp:Label>
        <asp:Label ID="D4CO4LEN" runat="server" Style="z-index: 128; left: 322px; position: absolute;
            top: 598px" Text=""></asp:Label>
        <asp:Label ID="D4CO3LEN" runat="server" Style="z-index: 129; left: 322px; position: absolute;
            top: 576px" Text=""></asp:Label>
        <asp:Label ID="D4CO2LEN" runat="server" Style="z-index: 130; left: 322px; position: absolute;
            top: 555px" Text=""></asp:Label>
        <asp:Label ID="D4CO5ITEM" runat="server" Style="z-index: 131; left: 245px; position: absolute;
            top: 620px" Text=""></asp:Label>
        <asp:Label ID="D4CO4ITEM" runat="server" Style="z-index: 132; left: 245px; position: absolute;
            top: 598px" Text=""></asp:Label>
        <asp:Label ID="D4CO3ITEM" runat="server" Style="z-index: 133; left: 245px; position: absolute;
            top: 576px" Text=""></asp:Label>
        <asp:Label ID="D4CO1LEN" runat="server" Style="z-index: 134; left: 322px; position: absolute;
            top: 532px" Text=""></asp:Label>
        <asp:Label ID="D4PILEN" runat="server" Style="z-index: 135; left: 322px; position: absolute;
            top: 510px" Text=""></asp:Label>
        <asp:Label ID="D4CO2ITEM" runat="server" Style="z-index: 136; left: 245px; position: absolute;
            top: 555px" Text=""></asp:Label>
        <asp:Label ID="D4PIITEM" runat="server" Style="z-index: 137; left: 245px; position: absolute;
            top: 510px" Text=""></asp:Label>
        <asp:Label ID="D4CO5COLNO" runat="server" Style="z-index: 138; left: 170px; position: absolute;
            top: 620px" Text=""></asp:Label>
        <asp:Label ID="D4CO4COLNO" runat="server" Style="z-index: 139; left: 170px; position: absolute;
            top: 598px" Text=""></asp:Label>
        <asp:Label ID="D4CO3COLNO" runat="server" Style="z-index: 140; left: 170px; position: absolute;
            top: 576px" Text=""></asp:Label>
        <asp:Label ID="D4CO2COLNO" runat="server" Style="z-index: 141; left: 170px; position: absolute;
            top: 555px" Text=""></asp:Label>
        <asp:Label ID="D4CO1COLNO" runat="server" Style="z-index: 142; left: 170px; position: absolute;
            top: 532px" Text=""></asp:Label>
        <asp:Label ID="D4CO5COL" runat="server" Style="z-index: 143; left: 94px; position: absolute;
            top: 620px" Text=""></asp:Label>
        <asp:Label ID="D4ECOL1" runat="server" Style="z-index: 144; left: 93px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4THRLOITEM" runat="server" Style="z-index: 145; left: 317px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="D4THRLOCOLNO" runat="server" Style="z-index: 146; left: 243px; position: absolute;
            top: 947px" Text=""></asp:Label>
        <asp:Label ID="D4THLLOITEM" runat="server" Style="z-index: 147; left: 317px; position: absolute;
            top: 928px" Text=""></asp:Label>
        <asp:Label ID="D4THLLOCOLNO" runat="server" Style="z-index: 148; left: 243px; position: absolute;
            top: 925px" Text=""></asp:Label>
        <asp:Label ID="D4THLOITEM" runat="server" Style="z-index: 149; left: 317px; position: absolute;
            top: 906px" Text=""></asp:Label>
        <asp:Label ID="D4THLOCOLNO" runat="server" Style="z-index: 150; left: 243px; position: absolute;
            top: 903px" Text=""></asp:Label>
        <asp:Label ID="D4THRUPITEM" runat="server" Style="z-index: 151; left: 317px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="D4THRUPCOLNO" runat="server" Style="z-index: 152; left: 243px; position: absolute;
            top: 880px" Text=""></asp:Label>
        <asp:Label ID="D4THLUPITEM" runat="server" Style="z-index: 153; left: 318px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4THLUPCOLNO" runat="server" Style="z-index: 154; left: 243px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4THUPITEM" runat="server" Style="z-index: 155; left: 317px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="D4THUPCOLNO" runat="server" Style="z-index: 156; left: 243px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="D4THRLOCOL" runat="server" Style="z-index: 157; left: 93px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="D4THLLOCOL" runat="server" Style="z-index: 158; left: 93px; position: absolute;
            top: 928px" Text=""></asp:Label>
        <asp:Label ID="D4THLOCOL" runat="server" Style="z-index: 159; left: 93px; position: absolute;
            top: 906px" Text=""></asp:Label>
        <asp:Label ID="D4THRUPCOL" runat="server" Style="z-index: 160; left: 93px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="D4THLUPCOL" runat="server" Style="z-index: 161; left: 93px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4THUPCOL" runat="server" Style="z-index: 162; left: 93px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="D4CCOL" runat="server" Style="z-index: 163; left: 93px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4CITEM" runat="server" Style="z-index: 164; left: 243px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4EITEM2" runat="server" Style="z-index: 165; left: 243px; position: absolute;
            top: 796px" Visible="False"></asp:Label>
        <asp:Label ID="D4EITEM1" runat="server" Style="z-index: 166; left: 243px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4TNRITEM2" runat="server" Style="z-index: 167; left: 319px; position: absolute;
            top: 752px" Visible="False"></asp:Label>
        <asp:Label ID="D4CNITEM" runat="server" Style="z-index: 168; left: 390px; position: absolute;
            top: 664px" Text=""></asp:Label>
        <asp:Label ID="D4CLTAUN" runat="server" Style="z-index: 169; left: 390px; position: absolute;
            top: 752px" Text=""></asp:Label>
        <asp:Label ID="D4COTAUN" runat="server" Style="z-index: 170; left: 539px; position: absolute;
            top: 752px" Text=""></asp:Label>
        <asp:Label ID="D4O2SUM" runat="server" Style="z-index: 171; left: 687px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="D4O2THRLOUN" runat="server" Style="z-index: 172; left: 687px; position: absolute;
            top: 951px" Text=""></asp:Label>
        <asp:Label ID="D4O2THLLOUN" runat="server" Style="z-index: 173; left: 686px; position: absolute;
            top: 929px" Text=""></asp:Label>
        <asp:Label ID="D4O1SUM" runat="server" Style="z-index: 174; left: 612px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="D4O1THRLOUN" runat="server" Style="z-index: 175; left: 612px; position: absolute;
            top: 951px" Text=""></asp:Label>
        <asp:Label ID="D4O1THLLOUN" runat="server" Style="z-index: 176; left: 611px; position: absolute;
            top: 929px" Text=""></asp:Label>
        <asp:Label ID="D4COTHRLOUN" runat="server" Style="z-index: 177; left: 540px; position: absolute;
            top: 951px"></asp:Label>
        <asp:Label ID="D4COUNIT" runat="server" Style="z-index: 178; left: 540px; position: absolute;
            top: 972px"></asp:Label>
        <asp:Label ID="D4COTHLLOUN" runat="server" Style="z-index: 179; left: 539px; position: absolute;
            top: 929px" Text=""></asp:Label>
        <asp:Label ID="D4O2THRUPUN" runat="server" Style="z-index: 180; left: 685px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="D4O2THLUPUN" runat="server" Style="z-index: 181; left: 685px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4O1THRUPUN" runat="server" Style="z-index: 182; left: 610px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="D4O1THLUPUN" runat="server" Style="z-index: 183; left: 610px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4COTHRUPUN" runat="server" Style="z-index: 184; left: 538px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="D4COTHLUPUN" runat="server" Style="z-index: 185; left: 538px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4O2CUN" runat="server" Style="z-index: 186; left: 688px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4O2MONOUN" runat="server" Style="z-index: 187; left: 687px; position: absolute;
            top: 796px" Text=""></asp:Label>
        <asp:Label ID="D4O1CUN" runat="server" Style="z-index: 188; left: 613px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4O1MONOUN" runat="server" Style="z-index: 189; left: 612px; position: absolute;
            top: 796px" Text=""></asp:Label>
        <asp:Label ID="D4COCUN" runat="server" Style="z-index: 190; left: 541px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4O2EUN" runat="server" Style="z-index: 191; left: 687px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4O1EUN" runat="server" Style="z-index: 192; left: 612px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4COEUN" runat="server" Style="z-index: 193; left: 540px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4O2TAUN" runat="server" Style="z-index: 194; left: 686px; position: absolute;
            top: 752px" Text=""></asp:Label>
        <asp:Label ID="D4O1TAUN" runat="server" Style="z-index: 196; left: 611px; position: absolute;
            top: 752px" Text=""></asp:Label>
        <asp:Label ID="D4CSUNIT" runat="server" Style="z-index: 197; left: 466px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="D4CSTHRLOUN" runat="server" Style="z-index: 198; left: 465px; position: absolute;
            top: 951px" Text=""></asp:Label>
        <asp:Label ID="D4CSTHLLOUN" runat="server" Style="z-index: 199; left: 465px; position: absolute;
            top: 929px" Text=""></asp:Label>
        <asp:Label ID="D4CLSUM" runat="server" Style="z-index: 200; left: 391px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="D4CLTHRLOUN" runat="server" Style="z-index: 201; left: 391px; position: absolute;
            top: 951px"></asp:Label>
        <asp:Label ID="D4CLTHLLOUN" runat="server" Style="z-index: 202; left: 391px; position: absolute;
            top: 929px" Text=""></asp:Label>
        <asp:Label ID="D4CSTHRUPUN" runat="server" Style="z-index: 203; left: 464px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="D4CSTHLUPUN" runat="server" Style="z-index: 204; left: 464px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4CLTHRUPUN" runat="server" Style="z-index: 205; left: 392px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="D4CLTHLUPUN" runat="server" Style="z-index: 206; left: 392px; position: absolute;
            top: 862px" Text=""></asp:Label>
        <asp:Label ID="D4CSCUN" runat="server" Style="z-index: 207; left: 467px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4CSEUN" runat="server" Style="z-index: 208; left: 466px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4CLCUN" runat="server" Style="z-index: 209; left: 392px; position: absolute;
            top: 818px" Text=""></asp:Label>
        <asp:Label ID="D4CLMONOUN" runat="server" Style="z-index: 210; left: 391px; position: absolute;
            top: 796px" Text=""></asp:Label>
        <asp:Label ID="D4CLEUN" runat="server" Style="z-index: 211; left: 391px; position: absolute;
            top: 774px" Text=""></asp:Label>
        <asp:Label ID="D4CSTAUN" runat="server" Style="z-index: 212; left: 465px; position: absolute;
            top: 752px" Text=""></asp:Label>
        <asp:Label ID="D4CSITEM" runat="server" Style="z-index: 213; left: 465px; position: absolute;
            top: 664px" Text=""></asp:Label>
        <asp:Label ID="D4CDITEM" runat="server" Style="z-index: 214; left: 539px; position: absolute;
            top: 664px" Text=""></asp:Label>
        <asp:Label ID="D4O1ITEM" runat="server" Style="z-index: 215; left: 611px; position: absolute;
            top: 664px" Text=""></asp:Label>
        <asp:Label ID="D4O2ITEM" runat="server" Style="z-index: 216; left: 686px; position: absolute;
            top: 664px" Text=""></asp:Label>
        <asp:Label ID="D4TNLITEM2" runat="server" Style="z-index: 218; left: 242px; position: absolute;
            top: 752px" Visible="False"></asp:Label>
        <asp:Label ID="D4ECOL2" runat="server" Style="z-index: 219; left: 93px; position: absolute;
            top: 796px" Visible="False"></asp:Label>
        <asp:Label ID="D4CO4COL" runat="server" Style="z-index: 220; left: 94px; position: absolute;
            top: 598px" Text=""></asp:Label>
        <asp:Label ID="D4CO3COL" runat="server" Style="z-index: 221; left: 94px; position: absolute;
            top: 576px" Text=""></asp:Label>
        <asp:Label ID="D4CO2COL" runat="server" Style="z-index: 222; left: 94px; position: absolute;
            top: 555px" Text=""></asp:Label>
        <asp:Label ID="D4CO1COL" runat="server" Style="z-index: 223; left: 94px; position: absolute;
            top: 532px" Text=""></asp:Label>
        <asp:Label ID="D4PICOLNO" runat="server" Style="z-index: 224; left: 170px; position: absolute;
            top: 510px" Text=""></asp:Label>
        <asp:Label ID="D4LYRLEN" runat="server" Style="z-index: 225; left: 658px; position: absolute;
            top: 400px" Text=""></asp:Label>
        <asp:Label ID="D4HMRLEN" runat="server" Style="z-index: 226; left: 658px; position: absolute;
            top: 378px" Text=""></asp:Label>
        <asp:Label ID="D4GMRLEN" runat="server" Style="z-index: 227; left: 658px; position: absolute;
            top: 356px" Text=""></asp:Label>
        <asp:Label ID="D4FMRLEN" runat="server" Style="z-index: 228; left: 658px; position: absolute;
            top: 334px" Text=""></asp:Label>
        <asp:Label ID="D4EMRLEN" runat="server" Style="z-index: 229; left: 658px; position: absolute;
            top: 312px" Text=""></asp:Label>
        <asp:Label ID="D4DMRLEN" runat="server" Style="z-index: 230; left: 658px; position: absolute;
            top: 290px" Text=""></asp:Label>
        <asp:Label ID="D4CMRLEN" runat="server" Style="z-index: 231; left: 658px; position: absolute;
            top: 268px" Text=""></asp:Label>
        <asp:Label ID="D4BMRLEN" runat="server" Style="z-index: 232; left: 658px; position: absolute;
            top: 246px" Text=""></asp:Label>
        <asp:Label ID="D4AMRLEN" runat="server" Style="z-index: 233; left: 658px; position: absolute;
            top: 224px" Text=""></asp:Label>
        <asp:Label ID="D4XMRLEN" runat="server" Style="z-index: 234; left: 658px; position: absolute;
            top: 202px" Text=""></asp:Label>
        <asp:Label ID="D4XMRITEM" runat="server" Style="z-index: 235; left: 579px; position: absolute;
            top: 202px"></asp:Label>
        <asp:Label ID="D4AMRITEM" runat="server" Style="z-index: 236; left: 579px; position: absolute;
            top: 224px"></asp:Label>
        <asp:Label ID="D4BMRITEM" runat="server" Style="z-index: 237; left: 579px; position: absolute;
            top: 246px"></asp:Label>
        <asp:Label ID="D4CMRITEM" runat="server" Style="z-index: 238; left: 579px; position: absolute;
            top: 268px"></asp:Label>
        <asp:Label ID="D4DMRITEM" runat="server" Style="z-index: 239; left: 579px; position: absolute;
            top: 290px"></asp:Label>
        <asp:Label ID="D4EMRITEM" runat="server" Style="z-index: 240; left: 579px; position: absolute;
            top: 312px"></asp:Label>
        <asp:Label ID="D4FMRITEM" runat="server" Style="z-index: 241; left: 579px; position: absolute;
            top: 334px"></asp:Label>
        <asp:Label ID="D4GMRITEM" runat="server" Style="z-index: 242; left: 579px; position: absolute;
            top: 356px"></asp:Label>
        <asp:Label ID="D4HMRITEM" runat="server" Style="z-index: 243; left: 579px; position: absolute;
            top: 378px"></asp:Label>
        <asp:Label ID="D4LYRITEM" runat="server" Style="z-index: 244; left: 579px; position: absolute;
            top: 400px"></asp:Label>
        <asp:Label ID="D4TCRITEM" runat="server" Style="z-index: 245; left: 579px; position: absolute;
            top: 422px"></asp:Label>
        <asp:Label ID="D4TNRITEM1" runat="server" Style="z-index: 246; left: 538px; position: absolute;
            top: 444px" Text=""></asp:Label>
        <asp:Label ID="D4XMRCOLNO" runat="server" Style="z-index: 247; left: 506px; position: absolute;
            top: 202px" Text=""></asp:Label>
        <asp:Label ID="D4AMRCOLNO" runat="server" Style="z-index: 248; left: 506px; position: absolute;
            top: 224px" Text=""></asp:Label>
        <asp:Label ID="D4BMRCOLNO" runat="server" Style="z-index: 249; left: 506px; position: absolute;
            top: 246px" Text=""></asp:Label>
        <asp:Label ID="D4CMRCOLNO" runat="server" Style="z-index: 250; left: 506px; position: absolute;
            top: 268px" Text=""></asp:Label>
        <asp:Label ID="D4DMRCOLNO" runat="server" Style="z-index: 251; left: 506px; position: absolute;
            top: 290px" Text=""></asp:Label>
        <asp:Label ID="D4EMRCOLNO" runat="server" Style="z-index: 252; left: 506px; position: absolute;
            top: 312px" Text=""></asp:Label>
        <asp:Label ID="D4FMRCOLNO" runat="server" Style="z-index: 253; left: 506px; position: absolute;
            top: 334px" Text=""></asp:Label>
        <asp:Label ID="D4GMRCOLNO" runat="server" Style="z-index: 254; left: 506px; position: absolute;
            top: 356px" Text=""></asp:Label>
        <asp:Label ID="D4HMRCOLNO" runat="server" Style="z-index: 255; left: 506px; position: absolute;
            top: 378px" Text=""></asp:Label>
        <asp:Label ID="D4LYRCOLNO" runat="server" Style="z-index: 256; left: 506px; position: absolute;
            top: 400px" Text=""></asp:Label>
        <asp:Label ID="D4TCRCOLNO" runat="server" Style="z-index: 257; left: 506px; position: absolute;
            top: 422px" Text=""></asp:Label>
        <asp:Label ID="D4XMRCOL" runat="server" Style="z-index: 258; left: 433px; position: absolute;
            top: 202px" Text=""></asp:Label>
        <asp:Label ID="D4AMRCOL" runat="server" Style="z-index: 259; left: 433px; position: absolute;
            top: 224px" Text=""></asp:Label>
        <asp:Label ID="D4BMRCOL" runat="server" Style="z-index: 260; left: 433px; position: absolute;
            top: 246px" Text=""></asp:Label>
        <asp:Label ID="D4CMRCOL" runat="server" Style="z-index: 261; left: 433px; position: absolute;
            top: 268px" Text=""></asp:Label>
        <asp:Label ID="D4DMRCOL" runat="server" Style="z-index: 262; left: 433px; position: absolute;
            top: 290px" Text=""></asp:Label>
        <asp:Label ID="D4EMRCOL" runat="server" Style="z-index: 263; left: 433px; position: absolute;
            top: 312px" Text=""></asp:Label>
        <asp:Label ID="D4FMRCOL" runat="server" Style="z-index: 264; left: 433px; position: absolute;
            top: 334px" Text=""></asp:Label>
        <asp:Label ID="D4GMRCOL" runat="server" Style="z-index: 265; left: 433px; position: absolute;
            top: 356px" Text=""></asp:Label>
        <asp:Label ID="D4HMRCOL" runat="server" Style="z-index: 266; left: 433px; position: absolute;
            top: 378px" Text=""></asp:Label>
        <asp:Label ID="D4LYRCOL" runat="server" Style="z-index: 267; left: 433px; position: absolute;
            top: 400px" Text=""></asp:Label>
        <asp:Label ID="D4TCRCOL" runat="server" Style="z-index: 268; left: 433px; position: absolute;
            top: 422px" Text=""></asp:Label>
        <asp:Label ID="D4XMLLEN" runat="server" Style="z-index: 269; left: 322px; position: absolute;
            top: 202px"></asp:Label>
        <asp:Label ID="D4AMLLEN" runat="server" Style="z-index: 270; left: 322px; position: absolute;
            top: 224px"></asp:Label>
        <asp:Label ID="D4BMLLEN" runat="server" Style="z-index: 271; left: 322px; position: absolute;
            top: 246px"></asp:Label>
        <asp:Label ID="D4CMLLEN" runat="server" Style="z-index: 272; left: 322px; position: absolute;
            top: 268px"></asp:Label>
        <asp:Label ID="D4DMLLEN" runat="server" Style="z-index: 273; left: 322px; position: absolute;
            top: 290px"></asp:Label>
        <asp:Label ID="D4EMLLEN" runat="server" Style="z-index: 274; left: 322px; position: absolute;
            top: 312px"></asp:Label>
        <asp:Label ID="D4FMLLEN" runat="server" Style="z-index: 275; left: 322px; position: absolute;
            top: 334px"></asp:Label>
        <asp:Label ID="D4GMLLEN" runat="server" Style="z-index: 276; left: 322px; position: absolute;
            top: 356px"></asp:Label>
        <asp:Label ID="D4HMLLEN" runat="server" Style="z-index: 277; left: 322px; position: absolute;
            top: 378px"></asp:Label>
        <asp:Label ID="D4LYLLEN" runat="server" Style="z-index: 278; left: 322px; position: absolute;
            top: 400px"></asp:Label>
        <asp:Label ID="D4TCLLEN" runat="server" Style="z-index: 279; left: 322px; position: absolute;
            top: 422px"></asp:Label>
        <asp:Label ID="D4YAMLSUM" runat="server" Style="z-index: 280; left: 317px; position: absolute;
            top: 444px" Text=""></asp:Label>
        <asp:Label ID="D4TNLITEM1" runat="server" Style="z-index: 281; left: 200px; position: absolute;
            top: 444px" Text=""></asp:Label>
        <asp:Label ID="D4XMLITEM" runat="server" Style="z-index: 282; left: 245px; position: absolute;
            top: 202px"></asp:Label>
        <asp:Label ID="D4AMLITEM" runat="server" Style="z-index: 283; left: 245px; position: absolute;
            top: 224px"></asp:Label>
        <asp:Label ID="D4BMLITEM" runat="server" Style="z-index: 284; left: 245px; position: absolute;
            top: 246px"></asp:Label>
        <asp:Label ID="D4CMLITEM" runat="server" Style="z-index: 285; left: 245px; position: absolute;
            top: 268px"></asp:Label>
        <asp:Label ID="D4DMLITEM" runat="server" Style="z-index: 286; left: 245px; position: absolute;
            top: 290px"></asp:Label>
        <asp:Label ID="D4EMLITEM" runat="server" Style="z-index: 287; left: 245px; position: absolute;
            top: 312px"></asp:Label>
        <asp:Label ID="D4FMLITEM" runat="server" Style="z-index: 288; left: 245px; position: absolute;
            top: 334px"></asp:Label>
        <asp:Label ID="D4GMLITEM" runat="server" Style="z-index: 289; left: 245px; position: absolute;
            top: 356px"></asp:Label>
        <asp:Label ID="D4HMLITEM" runat="server" Style="z-index: 290; left: 245px; position: absolute;
            top: 378px"></asp:Label>
        <asp:Label ID="D4LYLITEM" runat="server" Style="z-index: 291; left: 245px; position: absolute;
            top: 400px"></asp:Label>
        <asp:Label ID="D4TCLITEM" runat="server" Style="z-index: 292; left: 245px; position: absolute;
            top: 422px"></asp:Label>
        <asp:Label ID="D4XMLCOLNO" runat="server" Style="z-index: 293; left: 170px; position: absolute;
            top: 202px"></asp:Label>
        <asp:Label ID="D4AMLCOLNO" runat="server" Style="z-index: 294; left: 170px; position: absolute;
            top: 224px"></asp:Label>
        <asp:Label ID="D4BMLCOLNO" runat="server" Style="z-index: 295; left: 170px; position: absolute;
            top: 246px"></asp:Label>
        <asp:Label ID="D4CMLCOLNO" runat="server" Style="z-index: 296; left: 170px; position: absolute;
            top: 268px"></asp:Label>
        <asp:Label ID="D4DMLCOLNO" runat="server" Style="z-index: 297; left: 170px; position: absolute;
            top: 290px"></asp:Label>
        <asp:Label ID="D4EMLCOLNO" runat="server" Style="z-index: 298; left: 170px; position: absolute;
            top: 312px"></asp:Label>
        <asp:Label ID="D4FMLCOLNO" runat="server" Style="z-index: 299; left: 170px; position: absolute;
            top: 334px"></asp:Label>
        <asp:Label ID="D4GMLCOLNO" runat="server" Style="z-index: 300; left: 170px; position: absolute;
            top: 356px"></asp:Label>
        <asp:Label ID="D4HMLCOLNO" runat="server" Style="z-index: 301; left: 170px; position: absolute;
            top: 378px"></asp:Label>
        <asp:Label ID="D4LYLCOLNO" runat="server" Style="z-index: 302; left: 170px; position: absolute;
            top: 400px"></asp:Label>
        <asp:Label ID="D4TCLCOLNO" runat="server" Style="z-index: 303; left: 170px; position: absolute;
            top: 422px"></asp:Label>
        <asp:Label ID="D4AMLCOL" runat="server" Style="z-index: 304; left: 94px; position: absolute;
            top: 224px"></asp:Label>
        <asp:Label ID="D4BMLCOL" runat="server" Style="z-index: 305; left: 94px; position: absolute;
            top: 246px"></asp:Label>
        <asp:Label ID="D4CMLCOL" runat="server" Style="z-index: 306; left: 94px; position: absolute;
            top: 268px"></asp:Label>
        <asp:Label ID="D4DMLCOL" runat="server" Style="z-index: 307; left: 94px; position: absolute;
            top: 290px"></asp:Label>
        <asp:Label ID="D4EMLCOL" runat="server" Style="z-index: 308; left: 94px; position: absolute;
            top: 312px"></asp:Label>
        <asp:Label ID="D4FMLCOL" runat="server" Style="z-index: 309; left: 94px; position: absolute;
            top: 334px"></asp:Label>
        <asp:Label ID="D4GMLCOL" runat="server" Style="z-index: 310; left: 94px; position: absolute;
            top: 356px"></asp:Label>
        <asp:Label ID="D4HMLCOL" runat="server" Style="z-index: 311; left: 94px; position: absolute;
            top: 378px"></asp:Label>
        <asp:Label ID="D4LYLCOL" runat="server" Style="z-index: 312; left: 94px; position: absolute;
            top: 400px"></asp:Label>
        <asp:Label ID="D4TCLCOL" runat="server" Style="z-index: 313; left: 94px; position: absolute;
            top: 422px"></asp:Label>
        <asp:TextBox ID="D4OTHER2" runat="server" Height="38px" Style="z-index: 314; left: 685px;
            position: absolute; top: 687px;font-size:10px;background:transparent" TextMode="MultiLine" Width="69px" BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="D4OTHER1" runat="server" Height="38px" Style="z-index: 318; left: 612px;
            position: absolute; top: 687px;font-size:10px;background:transparent" TextMode="MultiLine" Width="69px" ReadOnly="True" BorderStyle="None" BorderWidth="0px"></asp:TextBox>
                    
                </table>
                    
              </asp:View>
            <!-- View-4 End -->
            <!-- -->
            <!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 核定履歷(View5) Start                     ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->
              <asp:View ID="View5" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">
            <asp:GridView style="Z-INDEX: 136; LEFT: 8px; POSITION: absolute; TOP: 1063px" id="GridView2" runat="server" Width="780px" Height="100px" BorderStyle="None" BackColor="White" BorderColor="#CC9966" CellPadding="4" BorderWidth="1px" AutoGenerateColumns="False">
            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099"  />
            <Columns>
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" ></asp:BoundField>
                <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" ></asp:BoundField>
                <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" ></asp:BoundField>
                <asp:BoundField DataField="StsDesc" HeaderText="處理結果" ></asp:BoundField>
                <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
                <asp:BoundField DataField="Description" HeaderText="核定時間">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                VerticalAlign="Middle"  />
        </asp:GridView>                
                </table>
              </asp:View>
            <!-- View-5 End -->
            <!-- -->
            </asp:MultiView>
        <!-- MultiView End -->
        <!-- -->
        <!-- +++++++++++++++++++++++++++++++++++++++++++++
             ++ 處理說明及Button Start                  ++ 
             +++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Button ID="BSAVE" runat="server" Height="23px" Style="z-index: 128; left: 416px;
            position: absolute; top: 1500px" Text="儲存" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG2" runat="server" Height="23px" Style="z-index: 129; left: 507px;
            position: absolute; top: 1500px" Text="NG2" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG1" runat="server" Height="23px" Style="z-index: 130; left: 599px;
            position: absolute; top: 1500px" Text="NG1" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BOK" runat="server" Height="23px" Style="z-index: 131; left: 691px;
            position: absolute; top: 1500px" Text="OK" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 132; left: 68px; position: absolute; top: 1425px"
            TextMode="MultiLine" Width="536px" Visible="False"></asp:TextBox>
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
            left: 22px; position: absolute; top: 1420px" />
        <!-- 處理說明及Button End -->
            <!-- -->
        <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox10" runat="server" Style="z-index: 126; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="不可為空白"></asp:RequiredFieldValidator>            
            
            
<!--*********************************************************************************************************** -->
    </div>
    </form>
</body>
</html>
