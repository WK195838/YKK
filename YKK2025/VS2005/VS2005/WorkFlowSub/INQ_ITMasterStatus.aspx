<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_ITMasterStatus.aspx.vb" Inherits="INQ_ITMasterStatus" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>IT MASTER Status</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox style="Z-INDEX: 126; LEFT: 16px; POSITION: absolute; TOP: 192px; TEXT-ALIGN: left" id="DMessqge1" runat="server" Width="288px" Height="18px" ForeColor="Blue" BorderStyle="Groove" BackColor="white" MaxLength="7">目前無委託在待處理中</asp:TextBox>

		<!-- ****************************************************************************************** -->
		<!-- ** System
		<!-- ****************************************************************************************** -->
        <img src="iMages/ITMaster_01.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 262px; height: 179px;"/>


        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 16px; position: absolute; top: 192px" Width="300px" ShowFooter="True">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <Columns>
                    <asp:BoundField DataField="FormNo" HeaderText="Form No" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="FormName" HeaderText="Form Name" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="RecCount" HeaderText="Count" ReadOnly="True">
                        <FooterStyle HorizontalAlign="RIGHT" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="RIGHT" />
                    </asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>

        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 350px; position: absolute; top: 8px" Width="500px" ShowFooter="True">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <Columns>
                    <asp:BoundField DataField="FormNo" HeaderText="Form No" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="FormName" HeaderText="Form Name" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="ReceiptTime" HeaderText="最早一筆時間" ReadOnly="True">
                        <FooterStyle HorizontalAlign="left" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="left" />
                    </asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>


        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 350px; position: absolute; top: 88px" Width="500px" ShowFooter="True">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <Columns>
                    <asp:BoundField DataField="FormNo" HeaderText="Form No" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="FormName" HeaderText="Form Name" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="ReceiptTime" HeaderText="最早一筆時間 Time" ReadOnly="True">
                        <FooterStyle HorizontalAlign="left" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="left" />
                    </asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>

        <asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 350px; position: absolute; top: 168px" Width="500px" ShowFooter="True">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <Columns>
                    <asp:BoundField DataField="FormNo" HeaderText="Form No" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="FormName" HeaderText="Form Name" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="ReceiptTime" HeaderText="最早一筆時間" ReadOnly="True">
                        <FooterStyle HorizontalAlign="left" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="left" />
                    </asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>

        <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 350px; position: absolute; top: 248px" Width="500px" ShowFooter="True">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <Columns>
                    <asp:BoundField DataField="FormNo" HeaderText="Form No" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="FormName" HeaderText="Form Name" ReadOnly="True">
                        <FooterStyle HorizontalAlign="LEFT" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="LEFT" />
                    </asp:BoundField>

                    <asp:BoundField DataField="ReceiptTime" HeaderText="最早一筆時間" ReadOnly="True">
                        <FooterStyle HorizontalAlign="left" />
                        <HeaderStyle Width="200px" />
                        <ItemStyle HorizontalAlign="left" />
                    </asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>


    </div>
    </form>
</body>

</html>
