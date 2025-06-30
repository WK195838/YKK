<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GentaniSheet.aspx.vb" Inherits="GentaniSheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>原單位</title>
    <style type="text/css">
    body { font-size:12px;}
    
    </style>
<script language="javascript" type="text/javascript">


// ]]>
</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        
    
     <img src="images/gentani_1.jpg" style="z-index: 1; left: 10px; position: absolute; top: 15px" />
     <img src="images/gentani_2.jpg" style="z-index: 1; left: 10px; position: absolute; top: 661px" />
     <asp:Label ID="DRNO" runat="server" Style="z-index: 100; left: 30px; position: absolute;
            top: 105px"></asp:Label>
        <asp:Label ID="DDEVNO" runat="server" Style="z-index: 101; left: 104px; position: absolute;
            top: 105px"></asp:Label>
        <asp:Label ID="DSIZENO" runat="server" Style="z-index: 102; left: 215px; position: absolute;
            top: 105px"></asp:Label>
        <asp:Label ID="DITEM" runat="server" Style="z-index: 103; left: 289px; position: absolute;
            top: 105px"></asp:Label>
        <asp:Label ID="DCODENO" runat="server" Style="z-index: 104; left: 360px; position: absolute;
            top: 105px"></asp:Label>
        <asp:Label ID="DMANUFTYPE" runat="server" Style="z-index: 105; left: 438px; position: absolute;
            top: 105px"></asp:Label>
        <asp:Label ID="DTALCOL" runat="server" Style="z-index: 106; left: 107px; position: absolute;
            top: 193px"></asp:Label>
        <asp:Label ID="DXMLCOL" runat="server" Style="z-index: 107; left: 107px; position: absolute;
            top: 215px"></asp:Label>
        <asp:Label ID="DTALCOLNO" runat="server" Style="z-index: 108; left: 183px; position: absolute;
            top: 193px"></asp:Label>
        <asp:Label ID="DTALITEM" runat="server" Style="z-index: 109; left: 258px; position: absolute;
            top: 193px"></asp:Label>
        <asp:Label ID="DTALLEN" runat="server" Style="z-index: 110; left: 335px; position: absolute;
            top: 193px"></asp:Label>
        <asp:Label ID="DTARCOL" runat="server" Style="z-index: 111; left: 446px; position: absolute;
            top: 193px" Text=""></asp:Label>
        <asp:Label ID="DTARCOLNO" runat="server" Style="z-index: 112; left: 519px; position: absolute;
            top: 193px" Text=""></asp:Label>
        <asp:Label ID="DTARITEM" runat="server" Style="z-index: 113; left: 592px; position: absolute;
            top: 193px"></asp:Label>
        <asp:Label ID="DTCRLEN" runat="server" Style="z-index: 114; left: 671px; position: absolute;
            top: 434px" Text=""></asp:Label>
        <asp:Label ID="DTARLEN" runat="server" Style="z-index: 115; left: 671px; position: absolute;
            top: 193px" Text=""></asp:Label>
        <asp:Label ID="DYAMRSUM" runat="server" Style="z-index: 116; left: 671px; position: absolute;
            top: 458px" Text=""></asp:Label>
        <asp:Label ID="DPICOL" runat="server" Style="z-index: 117; left: 107px; position: absolute;
            top: 523px" Text=""></asp:Label>
        <asp:Label ID="DCO1ITEM" runat="server" Style="z-index: 118; left: 258px; position: absolute;
            top: 545px" Text=""></asp:Label>
        <asp:Label ID="DTSLITEM" runat="server" Style="z-index: 119; left: 559px; position: absolute;
            top: 523px" Text=""></asp:Label>
        <asp:Label ID="DVMCORUNIT" runat="server" Style="z-index: 120; left: 669px; position: absolute;
            top: 632px" Text=""></asp:Label>
        <asp:Label ID="DTDRITEM" runat="server" Style="z-index: 121; left: 559px; position: absolute;
            top: 632px" Text=""></asp:Label>
        <asp:Label ID="DVMCOLUNIT" runat="server" Style="z-index: 122; left: 669px; position: absolute;
            top: 610px" Text=""></asp:Label>
        <asp:Label ID="DTDLITEM" runat="server" Style="z-index: 123; left: 559px; position: absolute;
            top: 610px" Text=""></asp:Label>
        <asp:Label ID="DVMSETRUNIT" runat="server" Style="z-index: 124; left: 669px; position: absolute;
            top: 545px" Text=""></asp:Label>
        <asp:Label ID="DVMSETLUNIT" runat="server" Style="z-index: 125; left: 669px; position: absolute;
            top: 523px" Text=""></asp:Label>
        <asp:Label ID="DTSRITEM" runat="server" Style="z-index: 126; left: 559px; position: absolute;
            top: 545px" Text=""></asp:Label>
        <asp:Label ID="DCO5LEN" runat="server" Style="z-index: 127; left: 335px; position: absolute;
            top: 633px" Text=""></asp:Label>
        <asp:Label ID="DCO4LEN" runat="server" Style="z-index: 128; left: 335px; position: absolute;
            top: 611px" Text=""></asp:Label>
        <asp:Label ID="DCO3LEN" runat="server" Style="z-index: 129; left: 335px; position: absolute;
            top: 589px" Text=""></asp:Label>
        <asp:Label ID="DCO2LEN" runat="server" Style="z-index: 130; left: 335px; position: absolute;
            top: 568px" Text=""></asp:Label>
        <asp:Label ID="DCO5ITEM" runat="server" Style="z-index: 131; left: 258px; position: absolute;
            top: 633px" Text=""></asp:Label>
        <asp:Label ID="DCO4ITEM" runat="server" Style="z-index: 132; left: 258px; position: absolute;
            top: 611px" Text=""></asp:Label>
        <asp:Label ID="DCO3ITEM" runat="server" Style="z-index: 133; left: 258px; position: absolute;
            top: 589px" Text=""></asp:Label>
        <asp:Label ID="DCO1LEN" runat="server" Style="z-index: 134; left: 335px; position: absolute;
            top: 545px" Text=""></asp:Label>
        <asp:Label ID="DPILEN" runat="server" Style="z-index: 135; left: 335px; position: absolute;
            top: 523px" Text=""></asp:Label>
        <asp:Label ID="DCO2ITEM" runat="server" Style="z-index: 136; left: 258px; position: absolute;
            top: 568px" Text=""></asp:Label>
        <asp:Label ID="DPIITEM" runat="server" Style="z-index: 137; left: 258px; position: absolute;
            top: 523px" Text=""></asp:Label>
        <asp:Label ID="DCO5COLNO" runat="server" Style="z-index: 138; left: 183px; position: absolute;
            top: 633px" Text=""></asp:Label>
        <asp:Label ID="DCO4COLNO" runat="server" Style="z-index: 139; left: 183px; position: absolute;
            top: 611px" Text=""></asp:Label>
        <asp:Label ID="DCO3COLNO" runat="server" Style="z-index: 140; left: 183px; position: absolute;
            top: 589px" Text=""></asp:Label>
        <asp:Label ID="DCO2COLNO" runat="server" Style="z-index: 141; left: 183px; position: absolute;
            top: 568px" Text=""></asp:Label>
        <asp:Label ID="DCO1COLNO" runat="server" Style="z-index: 142; left: 183px; position: absolute;
            top: 545px" Text=""></asp:Label>
        <asp:Label ID="DCO5COL" runat="server" Style="z-index: 143; left: 107px; position: absolute;
            top: 633px" Text=""></asp:Label>
        <asp:Label ID="DECOL1" runat="server" Style="z-index: 144; left: 107px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DTHRLOITEM" runat="server" Style="z-index: 145; left: 335px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="DTHRLOCOLNO" runat="server" Style="z-index: 146; left: 258px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="DTHLLOITEM" runat="server" Style="z-index: 147; left: 335px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DTHLLOCOLNO" runat="server" Style="z-index: 148; left: 258px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DTHLOITEM" runat="server" Style="z-index: 149; left: 335px; position: absolute;
            top: 928px" Text=""></asp:Label>
        <asp:Label ID="DTHLOCOLNO" runat="server" Style="z-index: 150; left: 258px; position: absolute;
            top: 928px" Text=""></asp:Label>
        <asp:Label ID="DTHRUPITEM" runat="server" Style="z-index: 151; left: 335px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DTHRUPCOLNO" runat="server" Style="z-index: 152; left: 258px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DTHLUPITEM" runat="server" Style="z-index: 153; left: 335px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DTHLUPCOLNO" runat="server" Style="z-index: 154; left: 258px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DTHUPITEM" runat="server" Style="z-index: 155; left: 335px; position: absolute;
            top: 861px" Text=""></asp:Label>
        <asp:Label ID="DTHUPCOLNO" runat="server" Style="z-index: 156; left: 258px; position: absolute;
            top: 861px" Text=""></asp:Label>
        <asp:Label ID="DTHRLOCOL" runat="server" Style="z-index: 157; left: 107px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="DTHLLOCOL" runat="server" Style="z-index: 158; left: 107px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DTHLOCOL" runat="server" Style="z-index: 159; left: 107px; position: absolute;
            top: 928px" Text=""></asp:Label>
        <asp:Label ID="DTHRUPCOL" runat="server" Style="z-index: 160; left: 107px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DTHLUPCOL" runat="server" Style="z-index: 161; left: 107px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DTHUPCOL" runat="server" Style="z-index: 162; left: 107px; position: absolute;
            top: 861px" Text=""></asp:Label>
        <asp:Label ID="DCCOL" runat="server" Style="z-index: 163; left: 107px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DCITEM" runat="server" Style="z-index: 164; left: 258px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DEITEM2" runat="server" Style="z-index: 165; left: 258px; position: absolute;
            top: 817px" Visible="False"></asp:Label>
        <asp:Label ID="DEITEM1" runat="server" Style="z-index: 166; left: 258px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DTNRITEM2" runat="server" Style="z-index: 167; left: 335px; position: absolute;
            top: 773px" Visible="False"></asp:Label>
        <asp:Label ID="DCNITEM" runat="server" Style="z-index: 168; left: 406px; position: absolute;
            top: 686px" Text=""></asp:Label>
        <asp:Label ID="DCLTAUN" runat="server" Style="z-index: 169; left: 406px; position: absolute;
            top: 773px" Text=""></asp:Label>
        <asp:Label ID="DCOTAUN" runat="server" Style="z-index: 170; left: 555px; position: absolute;
            top: 773px" Text=""></asp:Label>
        <asp:Label ID="DO2SUM" runat="server" Style="z-index: 171; left: 702px; position: absolute;
            top: 993px" Text=""></asp:Label>
        <asp:Label ID="DO2THRLOUN" runat="server" Style="z-index: 172; left: 702px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="DO2THLLOUN" runat="server" Style="z-index: 173; left: 702px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DO1SUM" runat="server" Style="z-index: 174; left: 627px; position: absolute;
            top: 993px" Text=""></asp:Label>
        <asp:Label ID="DO1THRLOUN" runat="server" Style="z-index: 175; left: 627px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="DO1THLLOUN" runat="server" Style="z-index: 176; left: 627px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DCOTHRLOUN" runat="server" Style="z-index: 177; left: 555px; position: absolute;
            top: 972px"></asp:Label>
        <asp:Label ID="DCOUNIT" runat="server" Style="z-index: 178; left: 555px; position: absolute;
            top: 993px"></asp:Label>
        <asp:Label ID="DCOTHLLOUN" runat="server" Style="z-index: 179; left: 555px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DO2THRUPUN" runat="server" Style="z-index: 180; left: 702px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DO2THLUPUN" runat="server" Style="z-index: 181; left: 702px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DO1THRUPUN" runat="server" Style="z-index: 182; left: 627px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DO1THLUPUN" runat="server" Style="z-index: 183; left: 627px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DCOTHRUPUN" runat="server" Style="z-index: 184; left: 555px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DCOTHLUPUN" runat="server" Style="z-index: 185; left: 555px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DO2CUN" runat="server" Style="z-index: 186; left: 702px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DO2MONOUN" runat="server" Style="z-index: 187; left: 702px; position: absolute;
            top: 817px" Text=""></asp:Label>
        <asp:Label ID="DO1CUN" runat="server" Style="z-index: 188; left: 627px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DO1MONOUN" runat="server" Style="z-index: 189; left: 627px; position: absolute;
            top: 817px" Text=""></asp:Label>
        <asp:Label ID="DCOCUN" runat="server" Style="z-index: 190; left: 555px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DO2EUN" runat="server" Style="z-index: 191; left: 702px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DO1EUN" runat="server" Style="z-index: 192; left: 627px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DCOEUN" runat="server" Style="z-index: 193; left: 555px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DO2TAUN" runat="server" Style="z-index: 194; left: 702px; position: absolute;
            top: 773px" Text=""></asp:Label>
        &nbsp;
        <asp:Label ID="DO1TAUN" runat="server" Style="z-index: 196; left: 627px; position: absolute;
            top: 773px" Text=""></asp:Label>
        <asp:Label ID="DCSUNIT" runat="server" Style="z-index: 197; left: 481px; position: absolute;
            top: 993px" Text=""></asp:Label>
        <asp:Label ID="DCSTHRLOUN" runat="server" Style="z-index: 198; left: 481px; position: absolute;
            top: 972px" Text=""></asp:Label>
        <asp:Label ID="DCSTHLLOUN" runat="server" Style="z-index: 199; left: 481px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DCLSUM" runat="server" Style="z-index: 200; left: 406px; position: absolute;
            top: 993px" Text=""></asp:Label>
        <asp:Label ID="DCLTHRLOUN" runat="server" Style="z-index: 201; left: 406px; position: absolute;
            top: 972px"></asp:Label>
        <asp:Label ID="DCLTHLLOUN" runat="server" Style="z-index: 202; left: 406px; position: absolute;
            top: 950px" Text=""></asp:Label>
        <asp:Label ID="DCSTHRUPUN" runat="server" Style="z-index: 203; left: 481px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DCSTHLUPUN" runat="server" Style="z-index: 204; left: 481px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DCLTHRUPUN" runat="server" Style="z-index: 205; left: 406px; position: absolute;
            top: 905px" Text=""></asp:Label>
        <asp:Label ID="DCLTHLUPUN" runat="server" Style="z-index: 206; left: 406px; position: absolute;
            top: 883px" Text=""></asp:Label>
        <asp:Label ID="DCSCUN" runat="server" Style="z-index: 207; left: 481px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DCSEUN" runat="server" Style="z-index: 208; left: 481px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DCLCUN" runat="server" Style="z-index: 209; left: 406px; position: absolute;
            top: 840px" Text=""></asp:Label>
        <asp:Label ID="DCLMONOUN" runat="server" Style="z-index: 210; left: 406px; position: absolute;
            top: 817px" Text=""></asp:Label>
        <asp:Label ID="DCLEUN" runat="server" Style="z-index: 211; left: 406px; position: absolute;
            top: 795px" Text=""></asp:Label>
        <asp:Label ID="DCSTAUN" runat="server" Style="z-index: 212; left: 481px; position: absolute;
            top: 773px" Text=""></asp:Label>
        <asp:Label ID="DCSITEM" runat="server" Style="z-index: 213; left: 481px; position: absolute;
            top: 686px" Text=""></asp:Label>
        <asp:Label ID="DCDITEM" runat="server" Style="z-index: 214; left: 555px; position: absolute;
            top: 686px" Text=""></asp:Label>
        <asp:Label ID="DO1ITEM" runat="server" Style="z-index: 215; left: 627px; position: absolute;
            top: 686px" Text=""></asp:Label>
        <asp:Label ID="DO2ITEM" runat="server" Style="z-index: 216; left: 702px; position: absolute;
            top: 686px" Text=""></asp:Label>
        &nbsp;
        <asp:Label ID="DTNLITEM2" runat="server" Style="z-index: 218; left: 258px; position: absolute;
            top: 773px" Visible="False"></asp:Label>
        <asp:Label ID="DECOL2" runat="server" Style="z-index: 219; left: 107px; position: absolute;
            top: 817px" Visible="False"></asp:Label>
        <asp:Label ID="DCO4COL" runat="server" Style="z-index: 220; left: 107px; position: absolute;
            top: 611px" Text=""></asp:Label>
        <asp:Label ID="DCO3COL" runat="server" Style="z-index: 221; left: 107px; position: absolute;
            top: 589px" Text=""></asp:Label>
        <asp:Label ID="DCO2COL" runat="server" Style="z-index: 222; left: 107px; position: absolute;
            top: 568px" Text=""></asp:Label>
        <asp:Label ID="DCO1COL" runat="server" Style="z-index: 223; left: 107px; position: absolute;
            top: 545px" Text=""></asp:Label>
        <asp:Label ID="DPICOLNO" runat="server" Style="z-index: 224; left: 183px; position: absolute;
            top: 523px" Text=""></asp:Label>
        <asp:Label ID="DLYRLEN" runat="server" Style="z-index: 225; left: 671px; position: absolute;
            top: 413px" Text=""></asp:Label>
        <asp:Label ID="DHMRLEN" runat="server" Style="z-index: 226; left: 671px; position: absolute;
            top: 390px" Text=""></asp:Label>
        <asp:Label ID="DGMRLEN" runat="server" Style="z-index: 227; left: 671px; position: absolute;
            top: 370px" Text=""></asp:Label>
        <asp:Label ID="DFMRLEN" runat="server" Style="z-index: 228; left: 671px; position: absolute;
            top: 347px" Text=""></asp:Label>
        <asp:Label ID="DEMRLEN" runat="server" Style="z-index: 229; left: 671px; position: absolute;
            top: 325px" Text=""></asp:Label>
        <asp:Label ID="DDMRLEN" runat="server" Style="z-index: 230; left: 671px; position: absolute;
            top: 303px" Text=""></asp:Label>
        <asp:Label ID="DCMRLEN" runat="server" Style="z-index: 231; left: 671px; position: absolute;
            top: 281px" Text=""></asp:Label>
        <asp:Label ID="DBMRLEN" runat="server" Style="z-index: 232; left: 671px; position: absolute;
            top: 259px" Text=""></asp:Label>
        <asp:Label ID="DAMRLEN" runat="server" Style="z-index: 233; left: 671px; position: absolute;
            top: 237px" Text=""></asp:Label>
        <asp:Label ID="DXMRLEN" runat="server" Style="z-index: 234; left: 671px; position: absolute;
            top: 215px" Text=""></asp:Label>
        <asp:Label ID="DXMRITEM" runat="server" Style="z-index: 235; left: 592px; position: absolute;
            top: 215px"></asp:Label>
        <asp:Label ID="DAMRITEM" runat="server" Style="z-index: 236; left: 592px; position: absolute;
            top: 237px"></asp:Label>
        <asp:Label ID="DBMRITEM" runat="server" Style="z-index: 237; left: 592px; position: absolute;
            top: 259px"></asp:Label>
        <asp:Label ID="DCMRITEM" runat="server" Style="z-index: 238; left: 592px; position: absolute;
            top: 281px"></asp:Label>
        <asp:Label ID="DDMRITEM" runat="server" Style="z-index: 239; left: 592px; position: absolute;
            top: 303px"></asp:Label>
        <asp:Label ID="DEMRITEM" runat="server" Style="z-index: 240; left: 592px; position: absolute;
            top: 325px"></asp:Label>
        <asp:Label ID="DFMRITEM" runat="server" Style="z-index: 241; left: 592px; position: absolute;
            top: 347px"></asp:Label>
        <asp:Label ID="DGMRITEM" runat="server" Style="z-index: 242; left: 592px; position: absolute;
            top: 370px"></asp:Label>
        <asp:Label ID="DHMRITEM" runat="server" Style="z-index: 243; left: 592px; position: absolute;
            top: 390px"></asp:Label>
        <asp:Label ID="DLYRITEM" runat="server" Style="z-index: 244; left: 592px; position: absolute;
            top: 413px"></asp:Label>
        <asp:Label ID="DTCRITEM" runat="server" Style="z-index: 245; left: 592px; position: absolute;
            top: 434px"></asp:Label>
        <asp:Label ID="DTNRITEM1" runat="server" Style="z-index: 246; left: 556px; position: absolute;
            top: 458px" Text=""></asp:Label>
        <asp:Label ID="DXMRCOLNO" runat="server" Style="z-index: 247; left: 519px; position: absolute;
            top: 215px" Text=""></asp:Label>
        <asp:Label ID="DAMRCOLNO" runat="server" Style="z-index: 248; left: 519px; position: absolute;
            top: 237px" Text=""></asp:Label>
        <asp:Label ID="DBMRCOLNO" runat="server" Style="z-index: 249; left: 519px; position: absolute;
            top: 259px" Text=""></asp:Label>
        <asp:Label ID="DCMRCOLNO" runat="server" Style="z-index: 250; left: 519px; position: absolute;
            top: 281px" Text=""></asp:Label>
        <asp:Label ID="DDMRCOLNO" runat="server" Style="z-index: 251; left: 519px; position: absolute;
            top: 303px" Text=""></asp:Label>
        <asp:Label ID="DEMRCOLNO" runat="server" Style="z-index: 252; left: 519px; position: absolute;
            top: 325px" Text=""></asp:Label>
        <asp:Label ID="DFMRCOLNO" runat="server" Style="z-index: 253; left: 519px; position: absolute;
            top: 347px" Text=""></asp:Label>
        <asp:Label ID="DGMRCOLNO" runat="server" Style="z-index: 254; left: 519px; position: absolute;
            top: 370px" Text=""></asp:Label>
        <asp:Label ID="DHMRCOLNO" runat="server" Style="z-index: 255; left: 519px; position: absolute;
            top: 390px" Text=""></asp:Label>
        <asp:Label ID="DLYRCOLNO" runat="server" Style="z-index: 256; left: 519px; position: absolute;
            top: 413px" Text=""></asp:Label>
        <asp:Label ID="DTCRCOLNO" runat="server" Style="z-index: 257; left: 519px; position: absolute;
            top: 434px" Text=""></asp:Label>
        <asp:Label ID="DXMRCOL" runat="server" Style="z-index: 258; left: 446px; position: absolute;
            top: 215px" Text=""></asp:Label>
        <asp:Label ID="DAMRCOL" runat="server" Style="z-index: 259; left: 446px; position: absolute;
            top: 237px" Text=""></asp:Label>
        <asp:Label ID="DBMRCOL" runat="server" Style="z-index: 260; left: 446px; position: absolute;
            top: 259px" Text=""></asp:Label>
        <asp:Label ID="DCMRCOL" runat="server" Style="z-index: 261; left: 446px; position: absolute;
            top: 281px" Text=""></asp:Label>
        <asp:Label ID="DDMRCOL" runat="server" Style="z-index: 262; left: 446px; position: absolute;
            top: 303px" Text=""></asp:Label>
        <asp:Label ID="DEMRCOL" runat="server" Style="z-index: 263; left: 446px; position: absolute;
            top: 325px" Text=""></asp:Label>
        <asp:Label ID="DFMRCOL" runat="server" Style="z-index: 264; left: 446px; position: absolute;
            top: 347px" Text=""></asp:Label>
        <asp:Label ID="DGMRCOL" runat="server" Style="z-index: 265; left: 446px; position: absolute;
            top: 370px" Text=""></asp:Label>
        <asp:Label ID="DHMRCOL" runat="server" Style="z-index: 266; left: 446px; position: absolute;
            top: 390px" Text=""></asp:Label>
        <asp:Label ID="DLYRCOL" runat="server" Style="z-index: 267; left: 446px; position: absolute;
            top: 413px" Text=""></asp:Label>
        <asp:Label ID="DTCRCOL" runat="server" Style="z-index: 268; left: 446px; position: absolute;
            top: 434px" Text=""></asp:Label>
        <asp:Label ID="DXMLLEN" runat="server" Style="z-index: 269; left: 335px; position: absolute;
            top: 215px"></asp:Label>
        <asp:Label ID="DAMLLEN" runat="server" Style="z-index: 270; left: 335px; position: absolute;
            top: 237px"></asp:Label>
        <asp:Label ID="DBMLLEN" runat="server" Style="z-index: 271; left: 335px; position: absolute;
            top: 259px"></asp:Label>
        <asp:Label ID="DCMLLEN" runat="server" Style="z-index: 272; left: 335px; position: absolute;
            top: 281px"></asp:Label>
        <asp:Label ID="DDMLLEN" runat="server" Style="z-index: 273; left: 335px; position: absolute;
            top: 303px"></asp:Label>
        <asp:Label ID="DEMLLEN" runat="server" Style="z-index: 274; left: 335px; position: absolute;
            top: 325px"></asp:Label>
        <asp:Label ID="DFMLLEN" runat="server" Style="z-index: 275; left: 335px; position: absolute;
            top: 347px"></asp:Label>
        <asp:Label ID="DGMLLEN" runat="server" Style="z-index: 276; left: 335px; position: absolute;
            top: 370px"></asp:Label>
        <asp:Label ID="DHMLLEN" runat="server" Style="z-index: 277; left: 335px; position: absolute;
            top: 390px"></asp:Label>
        <asp:Label ID="DLYLLEN" runat="server" Style="z-index: 278; left: 335px; position: absolute;
            top: 413px"></asp:Label>
        <asp:Label ID="DTCLLEN" runat="server" Style="z-index: 279; left: 335px; position: absolute;
            top: 434px"></asp:Label>
        <asp:Label ID="DYAMLSUM" runat="server" Style="z-index: 280; left: 335px; position: absolute;
            top: 458px" Text=""></asp:Label>
        <asp:Label ID="DTNLITEM1" runat="server" Style="z-index: 281; left: 218px; position: absolute;
            top: 458px" Text=""></asp:Label>
        <asp:Label ID="DXMLITEM" runat="server" Style="z-index: 282; left: 258px; position: absolute;
            top: 215px"></asp:Label>
        <asp:Label ID="DAMLITEM" runat="server" Style="z-index: 283; left: 258px; position: absolute;
            top: 237px"></asp:Label>
        <asp:Label ID="DBMLITEM" runat="server" Style="z-index: 284; left: 258px; position: absolute;
            top: 259px"></asp:Label>
        <asp:Label ID="DCMLITEM" runat="server" Style="z-index: 285; left: 258px; position: absolute;
            top: 281px"></asp:Label>
        <asp:Label ID="DDMLITEM" runat="server" Style="z-index: 286; left: 258px; position: absolute;
            top: 303px"></asp:Label>
        <asp:Label ID="DEMLITEM" runat="server" Style="z-index: 287; left: 258px; position: absolute;
            top: 325px"></asp:Label>
        <asp:Label ID="DFMLITEM" runat="server" Style="z-index: 288; left: 258px; position: absolute;
            top: 347px"></asp:Label>
        <asp:Label ID="DGMLITEM" runat="server" Style="z-index: 289; left: 258px; position: absolute;
            top: 370px"></asp:Label>
        <asp:Label ID="DHMLITEM" runat="server" Style="z-index: 290; left: 258px; position: absolute;
            top: 390px"></asp:Label>
        <asp:Label ID="DLYLITEM" runat="server" Style="z-index: 291; left: 258px; position: absolute;
            top: 413px"></asp:Label>
        <asp:Label ID="DTCLITEM" runat="server" Style="z-index: 292; left: 258px; position: absolute;
            top: 434px"></asp:Label>
        <asp:Label ID="DXMLCOLNO" runat="server" Style="z-index: 293; left: 183px; position: absolute;
            top: 215px"></asp:Label>
        <asp:Label ID="DAMLCOLNO" runat="server" Style="z-index: 294; left: 183px; position: absolute;
            top: 237px"></asp:Label>
        <asp:Label ID="DBMLCOLNO" runat="server" Style="z-index: 295; left: 183px; position: absolute;
            top: 259px"></asp:Label>
        <asp:Label ID="DCMLCOLNO" runat="server" Style="z-index: 296; left: 183px; position: absolute;
            top: 281px"></asp:Label>
        <asp:Label ID="DDMLCOLNO" runat="server" Style="z-index: 297; left: 183px; position: absolute;
            top: 303px"></asp:Label>
        <asp:Label ID="DEMLCOLNO" runat="server" Style="z-index: 298; left: 183px; position: absolute;
            top: 325px"></asp:Label>
        <asp:Label ID="DFMLCOLNO" runat="server" Style="z-index: 299; left: 183px; position: absolute;
            top: 347px"></asp:Label>
        <asp:Label ID="DGMLCOLNO" runat="server" Style="z-index: 300; left: 183px; position: absolute;
            top: 370px"></asp:Label>
        <asp:Label ID="DHMLCOLNO" runat="server" Style="z-index: 301; left: 183px; position: absolute;
            top: 390px"></asp:Label>
        <asp:Label ID="DLYLCOLNO" runat="server" Style="z-index: 302; left: 183px; position: absolute;
            top: 413px"></asp:Label>
        <asp:Label ID="DTCLCOLNO" runat="server" Style="z-index: 303; left: 183px; position: absolute;
            top: 434px"></asp:Label>
        <asp:Label ID="DAMLCOL" runat="server" Style="z-index: 304; left: 107px; position: absolute;
            top: 237px"></asp:Label>
        <asp:Label ID="DBMLCOL" runat="server" Style="z-index: 305; left: 107px; position: absolute;
            top: 259px"></asp:Label>
        <asp:Label ID="DCMLCOL" runat="server" Style="z-index: 306; left: 107px; position: absolute;
            top: 281px"></asp:Label>
        <asp:Label ID="DDMLCOL" runat="server" Style="z-index: 307; left: 107px; position: absolute;
            top: 303px"></asp:Label>
        <asp:Label ID="DEMLCOL" runat="server" Style="z-index: 308; left: 107px; position: absolute;
            top: 325px"></asp:Label>
        <asp:Label ID="DFMLCOL" runat="server" Style="z-index: 309; left: 107px; position: absolute;
            top: 347px"></asp:Label>
        <asp:Label ID="DGMLCOL" runat="server" Style="z-index: 310; left: 107px; position: absolute;
            top: 370px"></asp:Label>
        <asp:Label ID="DHMLCOL" runat="server" Style="z-index: 311; left: 107px; position: absolute;
            top: 390px"></asp:Label>
        <asp:Label ID="DLYLCOL" runat="server" Style="z-index: 312; left: 107px; position: absolute;
            top: 413px"></asp:Label>
        <asp:Label ID="DTCLCOL" runat="server" Style="z-index: 313; left: 107px; position: absolute;
            top: 434px"></asp:Label>
        <asp:TextBox ID="DOTHER2" runat="server" Height="38px" Style="z-index: 314; left: 688px;
            position: absolute; top: 705px;font-size:10px;background:transparent" TextMode="MultiLine" Width="69px" BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DOTHER1" runat="server" Height="38px" Style="z-index: 318; left: 615px;
            position: absolute; top: 705px;font-size:10px;background:transparent" TextMode="MultiLine" Width="69px" ReadOnly="True" BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        
    
       </div>
    </form>
</body>
</html>
