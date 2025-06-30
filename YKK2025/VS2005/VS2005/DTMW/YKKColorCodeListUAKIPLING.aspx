<%@ Page Language="VB" AutoEventWireup="false" CodeFile="YKKColorCodeListUAKIPLING.aspx.vb" Inherits="YKKColorCodeListUAKIPLING" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>色番資料</title>
    
    
    <script type="text/javascript">
  function ValidateNumber(e, pnumber) 
{
    if (!/^\d+[.]?[1-9]?$/.test(pnumber)) 
    {
        var newValue = /^\d+/.exec(e.value);
        
        if (newValue != null) 
        {  
            e.value = newValue;  
        }
        else 
        {  
            e.value = ""; 
        }
    }
    return false;
}
  
    </script>
    
</head>
<body>
    <form id="form1" runat="server">
        &nbsp;
      
      
  <asp:TextBox ID="DData" runat="server" BackColor="Yellow" ForeColor="Black" MaxLength="5"
               onkeyup="return ValidateNumber(this,value)" Style="z-index: 110; left: 136px;
               ime-mode: disabled; position: absolute; top: 13px" Width="51px"></asp:TextBox>
                  
               
        <asp:TextBox ID="DData1" runat="server" style="z-index: 100; left: 510px; position: absolute; top: 9px;text-transform : uppercase;" Height="17px" Width="55px" BackColor="Yellow" ForeColor="Black"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 579px; position: absolute; top: 10px" />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp;&nbsp;
   
 
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DData"
            ErrorMessage="請輸入Color Table" Style="left: -128px; position: relative; top: 53px; z-index: 103;" Width="274px"></asp:RequiredFieldValidator>&nbsp;<br />
        <br />
        &nbsp;&nbsp;
        <asp:Label ID="Label1" runat="server" Style="z-index: 104; left: 19px; position: absolute;
            top: 16px" Text="Color TABLE" ForeColor="Blue" Width="108px"></asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 105; left: 244px; position: absolute;
            top: 0px" Text="001 : MF / VFC   " ForeColor="#00C000"></asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 106; left: 243px; position: absolute;
            top: 24px" Text=" 002 : PF " ForeColor="#00C000"></asp:Label>
        <asp:Label ID="Label5" runat="server" Style="z-index: 107; left: 243px; position: absolute;
            top: 49px" Text="006 : VFO(M) / CIF / 3CFO" ForeColor="#00C000"></asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 108; left: 382px; position: absolute;
            top: 13px" Text="COLOR CODE" Width="115px" ForeColor="Blue"></asp:Label>
        <asp:Panel ID="Panel1" runat="server"  Height="300px" Width="593px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px" style="z-index: 109; left: 19px; position: absolute; top: 95px">     &nbsp;<asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="paccxx" HeaderText="Color Code"></asp:BoundField>
    <asp:BoundField DataField="cta1xx" HeaderText="左布帶" />
    <asp:BoundField DataField="ccp1xx" HeaderText="左鏈齒" />
    <asp:BoundField DataField="ccp2xx" HeaderText="右鏈齒" />
    <asp:BoundField DataField="cta2xx" HeaderText="右布帶" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> &nbsp; &nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
