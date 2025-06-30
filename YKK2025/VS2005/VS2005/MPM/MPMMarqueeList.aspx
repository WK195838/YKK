<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MPMMarqueeList.aspx.vb" Inherits="MPMMarqueeList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>跑馬燈記錄一覽</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="Black" BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 10px; position: absolute; top: 38px" Width="980px" DataKeyNames="Unique_ID">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:CommandField ButtonType="Button" ShowEditButton="True" EditText="變更" >
                    <ItemStyle BorderColor="#804040" BorderStyle="Groove"/>
                    <HeaderStyle BorderColor="Black" BorderStyle="Groove"/> 
                </asp:CommandField> 
                <asp:CommandField ButtonType="Button" ShowDeleteButton="True" >
                    <ItemStyle BorderColor="#804040" BorderStyle="Groove"/>
                    <HeaderStyle BorderColor="Black" BorderStyle="Groove"/> 
                </asp:CommandField> 
            <asp:BoundField HeaderText="狀態" DataField="Action" > 
                <ItemStyle HorizontalAlign="Center" BorderColor="#804040" BorderStyle="Groove"/>
                <HeaderStyle Width="80px" BorderColor="Black" BorderStyle="Groove"/> 
            </asp:BoundField>
            
            <asp:BoundField HeaderText="日期" DataField="AppDate" > 
                <ItemStyle HorizontalAlign="Center" BorderColor="#804040" BorderStyle="Groove"/>
                <HeaderStyle Width="100px" BorderColor="Black" BorderStyle="Groove"/> 
            </asp:BoundField>
            
            <asp:BoundField HeaderText="部門" DataField="Dep_Name" Visible="False" > 
                <ItemStyle HorizontalAlign="Center" BorderColor="#804040" BorderStyle="Groove"/>
                <HeaderStyle Width="80px" BorderColor="Black" BorderStyle="Groove"/> 
            </asp:BoundField>
            
            <asp:BoundField HeaderText="姓名" DataField="AppUser" >
                <ItemStyle HorizontalAlign="Center" BorderColor="#804040" BorderStyle="Groove"/>
                <HeaderStyle Width="80px" BorderColor="Black" BorderStyle="Groove"/> 
            </asp:BoundField>

            <asp:BoundField HeaderText="內容" DataField="Subject" >
                <ItemStyle HorizontalAlign="Left" BorderColor="#804040" BorderStyle="Groove"/>
                <HeaderStyle Width="400px" BorderColor="Black" BorderStyle="Groove"/> 
            </asp:BoundField>

            <asp:BoundField HeaderText="ID" DataField="Unique_ID" >  
            </asp:BoundField>

            </Columns>
        </asp:GridView>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:DropDownList ID="DSystem" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 102; left: 720px; position: absolute; top: 6px"
            Width="90px" Visible="False">
        </asp:DropDownList>
        <asp:Button ID="BGo" runat="server" Height="24px"
            Style="z-index: 110; left: 97px; position: absolute; top: 6px" Text="Go" Width="40px" CausesValidation="False" />
        <asp:DropDownList ID="DName" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="20px" Style="z-index: 119; left: 14px; position: absolute; top: 8px"
            Width="76px">
        </asp:DropDownList>
        <asp:Button ID="BNewItem" runat="server" Height="24px"
            Style="z-index: 110; left: 141px; position: absolute; top: 7px" Text="新工作記錄" Width="95px" />
        <asp:DropDownList ID="DType" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 102; left: 628px; position: absolute; top: 7px"
            Width="90px" Visible="False">
        </asp:DropDownList><asp:Button ID="Button1" runat="server" Height="24px"
            Style="z-index: 110; left: 526px; position: absolute; top: 8px" Text="回工作週報" Width="95px" Visible="False" />
    </div>
    </form></body>
</html>
