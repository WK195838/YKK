<%@ Page Language="VB" AutoEventWireup="true" EnableEventValidation = "false" CodeFile="FindItemPage.aspx.vb" Inherits="FindItemPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >


<head id="Head1" runat="server">
    <title>Wave's Item搜尋(AS400)</title>>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 101px; position: absolute;
            top: 5px; text-align: left" Width="98px" MaxLength="7"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 6px; text-align: left" Width="57px">CODE</asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 32px; text-align: left" Width="88px">NAME-1</asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 55px; text-align: left" Width="88px">AND NAME-2</asp:TextBox>
        <asp:TextBox ID="TextBox5" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 80px; text-align: left" Width="88px">AND NAME-3</asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 13px; position: absolute;
            top: 105px; text-align: left" Width="88px">AND NAME-4</asp:TextBox>
        <asp:TextBox ID="DName1" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 101px; position: absolute; top: 29px;
            text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DName2" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 126; left: 101px; position: absolute; top: 54px;
            text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DName3" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" MaxLength="35" Style="z-index: 126; left: 101px; position: absolute;
            top: 79px; text-align: left" Width="289px"></asp:TextBox>
        <asp:Button ID="BFindItem" runat="server" Height="25px" Style="z-index: 104; left: 399px;
            position: absolute; top: 103px" Text="搜尋" Width="45px" />

		<asp:hyperlink id="LShortSheet" style="Z-INDEX: 102; LEFT: 584px; POSITION: absolute; TOP: 8px" runat="server"
					Width="130px" Height="16px" Font-Size="12pt" NavigateUrl="https://ykkglobal.sharepoint.com/sites/asia_twn_discuss_mktsal/Lists/IRW1/DispForm.aspx?ID=34&e=ysgi5n&xsdata=MDV8MDJ8fDZhYTg0OGFjZGM0MTQ4MTYxMjFkMDhkY2ZlZGQwYzAzfDUxYTJlMTYzZDNlMTQwYTZiODI1MjZhNGIwYjM4MzEwfDB8MHw2Mzg2NjU0NzI1MjI1MjMxNjV8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPakE0WkdNNE5qVTBMV1EwT1RBdE5EZzBNQzA1TlRNeExXUXhNRFl5TlRVM1lqWmxObDh4TjJZNE56SXdOeTB5TW1SbExUUm1PVFV0WW1KaU1DMDVZemt6T0RZek1EZ3dOMkpBZFc1eExtZGliQzV6Y0dGalpYTXZiV1Z6YzJGblpYTXZNVGN6TURrMU1EUTFNVEEwTWc9PXxiNDE0OGZkZTM5YzQ0ZGQxMTIxZDA4ZGNmZWRkMGMwM3xjYTllMzlkMmRmYTI0YjJiOTYyNDEzZDRiNmRjNWMyMg%3D%3D&sdata=RndxanpETEdWazRuNWVnNTFnMldMb2FPN0ZlVGw3Tzk0T09ObE1GV2RSZz0%3D&ovuser=51a2e163-d3e1-40a6-b825-26a4b0b38310%2Cjoo%40ykk.com&OR=Teams-HL&CT=1730950453142&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yNDEwMDMyMDkwMCIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D" Target="_blank">拉頭品名字元限制</asp:hyperlink>


        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 456px; position: absolute; top: 104px" Width="24px" />


        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="25" Style="z-index: 114; left: 12px; position: absolute; top: 131px"
            Width="1200px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:CommandField ShowSelectButton="True" SelectText="報價" />

                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="ISOS2FAS" HeaderText="" Target="_blank">
                </asp:HyperLinkField>
                
                <asp:BoundField DataField="Code" HeaderText="Item Code" />
                <asp:BoundField DataField="Name1" HeaderText="Item Name-1" />
                <asp:BoundField DataField="Name2" HeaderText="Item Name-2" />
                <asp:BoundField DataField="Name3" HeaderText="Item Name-3" />
                <asp:BoundField DataField="Slider" HeaderText="Slider Code" />
                <asp:BoundField DataField="Season" HeaderText="Season" />
                <asp:BoundField DataField="Color" HeaderText="Color" />
                <asp:BoundField DataField="ColorName" HeaderText="Color Name" />

                <asp:HyperLinkField DataNavigateUrlFields="BCPURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="BCP" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="BCTURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="BCT" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="IRWURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="IRW" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SPCURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SPC" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SPPURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SPP" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="IMGURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="IMG" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="COMBIURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="COMBI" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="VALURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="VAL" HeaderText="" Target="_blank">
                </asp:HyperLinkField>

            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" BorderStyle="Groove" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
        <asp:TextBox ID="DName4" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="35" Style="z-index: 126; left: 101px;
            position: absolute; top: 104px; text-align: left" Width="289px"></asp:TextBox>

    </div>
    </form>
</body>

</html>
