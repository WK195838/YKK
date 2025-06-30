<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GentaniSheet.aspx.vb" Inherits="GentaniSheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
                    <asp:Image ID="Image3" runat="server" ImageUrl="~/Images/DevelopmentGentani_01.png" Style="z-index: 99;
                    left: 5px; position: absolute; top: 2px" />
                    
     <asp:Label ID="D4NO" runat="server" Style="z-index: 100; left: 21px; position: absolute;
            top: 92px" Width="70px"></asp:Label>
        <asp:Label ID="D4DEVNO" runat="server" Style="z-index: 101; left: 205px; position: absolute;
            top: 92px" Width="106px"></asp:Label>
        <asp:Label ID="D4SIZENO" runat="server" Style="z-index: 102; left: 316px; position: absolute;
            top: 92px" Width="9px"></asp:Label>
        <asp:Label ID="D4ITEM" runat="server" Style="z-index: 103; left: 390px; position: absolute;
            top: 92px"></asp:Label>
        <asp:Label ID="D4CODENO" runat="server" Style="z-index: 104; left: 463px; position: absolute;
            top: 92px"></asp:Label>
        <asp:Label ID="D4MANUFTYPE" runat="server" Style="z-index: 105; left: 538px; position: absolute;
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
            position: absolute; top: 687px;font-size:10px;background:transparent" TextMode="MultiLine" Width="69px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
                    
    
    </div>
    </form>
</body>
</html>
