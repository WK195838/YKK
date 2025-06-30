<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AdvencedImages.aspx.vb" Inherits="AdvencedImages" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Advenced Images</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ****************************************************************************************** -->
        <!-- ** Button                                                                                  -->
        <!-- ****************************************************************************************** -->
        <!-- R&D Find -->
                <asp:Button ID="BFind" runat="server" Height="64px" Style="z-index: 104; left: 304px;
                    position: absolute; top: 16px" Text="R&D" Width="96px" Font-Bold="True" Font-Size="Larger" />
        <!-- Wings Find -->
                <asp:Button ID="BFindWings" runat="server" Height="64px" Style="z-index: 104; left: 416px;
                    position: absolute; top: 16px" Text="WINGS" Width="96px" Font-Bold="True" Font-Size="Larger" />
        &nbsp;<!-- Excel -->

		<!-- ****************************************************************************************** -->
		<!-- ** Puller Key                                                                              -->
		<!-- ****************************************************************************************** -->
            <asp:TextBox ID="TextBox2" runat="server" BackColor="WHITE" BorderStyle="None" ForeColor="BLACK"
                Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
                top: 16px; text-align: left" Width="288px" Font-Bold="True" Font-Size="Larger">Sample: Series=Y, TW, AD, NK</asp:TextBox>
        &nbsp; &nbsp;&nbsp;

            <asp:TextBox ID="TextBox3" runat="server" BackColor="Black" BorderStyle="None" Font-Bold="True"
                Font-Size="Larger" ForeColor="White" Height="20px" ReadOnly="True" Style="z-index: 126;
                left: 8px; position: absolute; top: 48px; text-align: left" Width="120px">Series</asp:TextBox>

            <asp:TextBox ID="DKSeries" runat="server" BackColor="Yellow" BorderStyle="Groove"
                Font-Bold="False" Font-Size="Larger" ForeColor="Blue" Height="18px" Style="z-index: 126;
                left: 128px; position: absolute; top: 48px; text-align: left" Width="128px"></asp:TextBox>


            <asp:TextBox ID="DStatus" runat="server" BackColor="white" BorderStyle="None" ForeColor="black"
                Height="56px" ReadOnly="True" Style="z-index: 126; left:-528px; position: absolute;
                top: 16px; text-align: left" Width="416px" Font-Bold="false" Font-Size="Medium" TextMode="MultiLine"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- ** DataList                                                                          -->
		<!-- ****************************************************************************************** -->
			<asp:datalist id="DataList1" style="Z-INDEX: 105; POSITION: absolute; TOP: 88px; LEFT: 8px" runat="server"
					Width="10px" Height="93px" CellSpacing="2" RepeatColumns="10" RepeatDirection="Horizontal" BorderColor="Black" BorderStyle="Groove" BorderWidth="3">
					<ItemTemplate>
						<asp:Table Width="100%" GridLines="Both" Font-Size="8pt" Runat="server" ID="Table2">
							<asp:TableRow>
								<asp:TableCell HorizontalAlign="Left" BorderWidth="1" BorderStyle="Solid" BorderColor="#990000"
									Font-Size="10" BackColor="Black" ForeColor="White">
									<%#Container.DataItem("Code")%>
                                </asp:TableCell>
							</asp:TableRow>
							<asp:TableRow>
								<asp:TableCell HorizontalAlign="Left" BorderWidth="1" BorderStyle="Solid" BorderColor="Black"
									Font-Size="10" BackColor="Yellow">

                                    <asp:HyperLink ID="HyperLink1" Text='<%# Container.DataItem("No")%>' NavigateUrl='<%# Container.DataItem("NoURL")%>'  runat="server" Target="_blank"></asp:HyperLink>

								</asp:TableCell>
							</asp:TableRow>
							<asp:TableRow>
								<asp:TableCell BorderWidth="1" BorderStyle="Solid" BorderColor="Black" HorizontalAlign="Center"
									BackColor="Black" >
									<asp:Image ID="Image1" Runat="server" Height="186" Width="163" ImageUrl='<%# Container.DataItem("ImagePath") %>' >
									</asp:Image>
								</asp:TableCell>
							</asp:TableRow>
						</asp:Table>
					</ItemTemplate>
				</asp:datalist>
    </div>
    </form>
</body>
</html>
