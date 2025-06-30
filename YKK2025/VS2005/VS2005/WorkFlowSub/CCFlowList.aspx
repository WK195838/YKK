<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CCFlowList.aspx.vb" Inherits="CCFlowList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>CC Flow List</title>
  	<script language="javascript" type="text/javascript">
	</script>
</head>
<body>
  	<form id="Form1" runat="server">
  	 <div>
           <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
               BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
               Style="z-index: 104; left: 8px; position: absolute; top: 10px" Width="1080px">
               <Columns>
                   <asp:BoundField DataField="FormName" HeaderText="委託單" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="No" HeaderText="No." >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="ApplyName" HeaderText="委託人" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="StepNameDesc" HeaderText="工程" >
                       <ItemStyle Width="140px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="OP" HeaderText="" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>

                   <asp:BoundField DataField="URL" HeaderText="xURL" />
                   <asp:BoundField DataField="OPURL" HeaderText="xOPURL" />

               </Columns>
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:GridView>
  	 
      </div>
    </form>
</body>
</html>
