<%@ Page Language="vb" AutoEventWireup="false" Inherits="DispoalSendDate" aspCompat="True" EnableEventValidation = "false"  CodeFile="DispoalSendDate.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>報廢處理申請書調閱資料</title>
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DisposalSendTime_02.jpg" style="z-index: 99; left: 10px; position: absolute; top: 15px" /><asp:DropDownList ID="DAccount1" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 522px; position: absolute; top: 179px" Width="105px">
            <asp:ListItem></asp:ListItem>
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
            <asp:ListItem>03</asp:ListItem>
            <asp:ListItem>04</asp:ListItem>
            <asp:ListItem>05</asp:ListItem>
            <asp:ListItem>06</asp:ListItem>
            <asp:ListItem>07</asp:ListItem>
            <asp:ListItem>08</asp:ListItem>
            <asp:ListItem>09</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>11</asp:ListItem>
            <asp:ListItem>12</asp:ListItem>
            <asp:ListItem>13</asp:ListItem>
            <asp:ListItem>14</asp:ListItem>
            <asp:ListItem>15 </asp:ListItem>
            <asp:ListItem>16</asp:ListItem>
            <asp:ListItem>17</asp:ListItem>
            <asp:ListItem>18</asp:ListItem>
            <asp:ListItem>19</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>21</asp:ListItem>
            <asp:ListItem>22</asp:ListItem>
            <asp:ListItem>23</asp:ListItem>
            <asp:ListItem>24</asp:ListItem>
            <asp:ListItem>25</asp:ListItem>
            <asp:ListItem>26</asp:ListItem>
            <asp:ListItem>27</asp:ListItem>
            <asp:ListItem>28</asp:ListItem>
            <asp:ListItem>29</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
            <asp:ListItem>31</asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="BPCE1" runat="server" CausesValidation="False" Height="21px" Style="z-index: 101;
            left: 346px; position: absolute; top: 180px" Text="....." Width="20px" />
        <asp:TextBox ID="DPCE1" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 102; left: 259px; position: absolute;
            top: 180px" Width="84px"></asp:TextBox><asp:Button ID="BPCE2" runat="server" CausesValidation="False" Height="21px" Style="z-index: 101;
            left: 346px; position: absolute; top: 224px" Text="....." Width="20px" />
        <asp:TextBox ID="DPCE2" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 102; left: 259px; position: absolute;
            top: 222px" Width="84px"></asp:TextBox>
        <asp:Button ID="BPCE3" runat="server" CausesValidation="False" Height="21px" Style="z-index: 101;
            left: 347px; position: absolute; top: 269px" Text="....." Width="20px" />
        <asp:TextBox ID="DPCE3" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 102; left: 259px; position: absolute;
            top: 268px" Width="84px"></asp:TextBox>
        <asp:DropDownList ID="DPCS1" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 113px; position: absolute; top: 180px" Width="105px">
            <asp:ListItem></asp:ListItem>
          <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>11</asp:ListItem>
            <asp:ListItem>12</asp:ListItem>
            <asp:ListItem>13</asp:ListItem>
            <asp:ListItem>14</asp:ListItem>
            <asp:ListItem>15 </asp:ListItem>
            <asp:ListItem>16</asp:ListItem>
            <asp:ListItem>17</asp:ListItem>
            <asp:ListItem>18</asp:ListItem>
            <asp:ListItem>19</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>21</asp:ListItem>
            <asp:ListItem>22</asp:ListItem>
            <asp:ListItem>23</asp:ListItem>
            <asp:ListItem>24</asp:ListItem>
            <asp:ListItem>25</asp:ListItem>
            <asp:ListItem>26</asp:ListItem>
            <asp:ListItem>27</asp:ListItem>
            <asp:ListItem>28</asp:ListItem>
            <asp:ListItem>29</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
            <asp:ListItem>31</asp:ListItem>
        </asp:DropDownList>
          <asp:DropDownList ID="DPCS3" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 115px; position: absolute; top: 268px" Width="105px">
              <asp:ListItem></asp:ListItem>
              <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
          <asp:DropDownList ID="DAdmin1" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 397px; position: absolute; top: 178px" Width="105px">
              <asp:ListItem></asp:ListItem>
               <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
          <asp:DropDownList ID="DAdmin2" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 397px; position: absolute; top: 222px" Width="105px">
              <asp:ListItem></asp:ListItem>
             <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
          <asp:DropDownList ID="DAdmin3" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 397px; position: absolute; top: 267px" Width="105px">
              <asp:ListItem></asp:ListItem>
             <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
          <asp:DropDownList ID="DPCS2" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 114px; position: absolute; top: 222px" Width="105px">
              <asp:ListItem></asp:ListItem>
              <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
          &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:Button ID="BSetDate" runat="server" Style="z-index: 109; left: 564px; position: absolute;
            top: 312px" Text="設定" Width="62px" />
        <asp:Label ID="LYM1" runat="server" Height="3px" Style="z-index: 110; left: 30px;
            position: absolute; top: 178px" Text="2015/01"></asp:Label>
        <asp:Label ID="LYM2" runat="server" Height="3px" Style="z-index: 111; left: 31px;
            position: absolute; top: 219px" Text="2015/01"></asp:Label>
        <asp:Label ID="LYM3" runat="server" Height="3px" Style="z-index: 112; left: 31px;
            position: absolute; top: 265px" Text="2015/01"></asp:Label>
          &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
          <asp:DropDownList ID="DAccount2" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 522px; position: absolute; top: 222px" Width="105px">
              <asp:ListItem></asp:ListItem>
              <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
          <asp:DropDownList ID="DAccount3" runat="server" AutoPostBack="True" BackColor="GreenYellow"
            Style="z-index: 103; left: 523px; position: absolute; top: 266px" Width="105px">
              <asp:ListItem></asp:ListItem>
              <asp:ListItem>01</asp:ListItem>
              <asp:ListItem>02</asp:ListItem>
              <asp:ListItem>03</asp:ListItem>
              <asp:ListItem>04</asp:ListItem>
              <asp:ListItem>05</asp:ListItem>
              <asp:ListItem>06</asp:ListItem>
              <asp:ListItem>07</asp:ListItem>
              <asp:ListItem>08</asp:ListItem>
              <asp:ListItem>09</asp:ListItem>
              <asp:ListItem>10</asp:ListItem>
              <asp:ListItem>11</asp:ListItem>
              <asp:ListItem>12</asp:ListItem>
              <asp:ListItem>13</asp:ListItem>
              <asp:ListItem>14</asp:ListItem>
              <asp:ListItem>15 </asp:ListItem>
              <asp:ListItem>16</asp:ListItem>
              <asp:ListItem>17</asp:ListItem>
              <asp:ListItem>18</asp:ListItem>
              <asp:ListItem>19</asp:ListItem>
              <asp:ListItem>20</asp:ListItem>
              <asp:ListItem>21</asp:ListItem>
              <asp:ListItem>22</asp:ListItem>
              <asp:ListItem>23</asp:ListItem>
              <asp:ListItem>24</asp:ListItem>
              <asp:ListItem>25</asp:ListItem>
              <asp:ListItem>26</asp:ListItem>
              <asp:ListItem>27</asp:ListItem>
              <asp:ListItem>28</asp:ListItem>
              <asp:ListItem>29</asp:ListItem>
              <asp:ListItem>30</asp:ListItem>
              <asp:ListItem>31</asp:ListItem>
          </asp:DropDownList>
    </div>
    </form>
</body>
</html>
