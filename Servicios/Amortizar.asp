<% Response.Expires=0 
   Response.Buffer = True

   Cta_Credito = Request("Cta_Credito")   
   Interes = CDbl(Request("Interes"))
   Importe = CDbl(Request("Importe")) + Interes
   Tope = Request("Tope")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Amortización de Préstamos</title>
<style type="text/css">
<!-- 
.style1 {
	font-family: Verdana;
	font-size: 12px;
}
.style3 {
	font-family: Verdana;
	color: #000000;
	font-weight: bold;
	font-size: 16px;
	
}
.style11 {
	font-family: Verdana;
	font-weight: bold;
	color: #FFFFFF;
	font-size: 20px; 
	filter:glow(color=#000000,strength=2);
    width:100%;
}
-->
</style>
 <style>
 
 body {
	scrollbar-arrow-color: #DDCE67;
	scrollbar-base-color: #f0b6c4;
	scrollbar-face-color: #9f314e;
	scrollbar-highlight-color: #f0b6c4;
	scrollbar-shadow-color: #DDCE67;
}
 </style>

<style>
		A		{
			text-decoration : none;
			color : "#0000FF";
			font-size : 14px;		}

		A:hover		{
			text-decoration : underline;
			color : "#FF0000";		}
	.style12 {
	font-family: Verdana;
	font-size: 18px;
	font-weight: bold;
	color: #FF0000;
}
.style13 {
	font-family: Verdana;
	font-size: 16px;
	font-weight: bold;
	color: #FF0000;
}
.style14 {font-family: Verdana}
</style>
<script language="JavaScript">

function solo_num(theComp)
{
  var checkOK = "0123456789.";
  var checkStr = theComp.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch == ".")
    {
      decPoints++;
      if ( decPoints == 1 ) allNum += ".";
    }
    else
      allNum += ch;
  }
  
  if (!allValid)
  {
    alert("Por favor escriba solo dígitos en este campo. El separador decimal permitido es el punto. No necesita separar los miles.");
    theComp.value = allNum
    theComp.focus();
    return (false);
  }

  if (decPoints > 1)
  {
    //alert("Por favor escriba un numero válido en este campo \"Importe\". El separador decimal permitido es el punto. No necesita separar los miles.");
    theComp.value = allNum
    theComp.focus();
    return (false);
  }
  
  return (true);
}

</script>
</head>
<!-- #Include file = "Informa.asp"-->

<body style="background-color: transparent " topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style11">Amortización de Fondos.</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
</table>

<table width="100%" border="0" Align="Center" cellpadding="5" cellspacing="5">
  <tr>
    <td height="30" align="right" bgcolor="#FCE8AB">
      <span class="style3">Fecha:&nbsp;</span><span class="style3"><%=Date%></span></td>
  </tr>
</table>

<% 
  sucu = mid(Session("UsrId"), 1 , 4)
  whois = mid(Session("UsrId"), 5 , 4)

  Application.Lock 
  Application("NumVisits") = Application("NumVisits") + 1 
  Application.Unlock 

  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")

  file = "CUE" & sucu
  query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM "& file &" WHERE (Cod_Contra = '" & whois & "') AND (Cue_Inact = .F.) AND (Cue_Cierre = .F.) AND (Cue_Cierre = .F.) AND (CUE_SUBCUE = '3210') AND (SIG_MONEDA = '"& mid(Cta_Credito,1,3) &"')"
  set rs = conn.Execute( query )

  On Error Resume Next
  If Request.Form("hname") = "" Then
  
  ' Permite la captación de datos en la forma HTML
  
  Function Resto (cuenta)
     Cuenta=TRIM(Cuenta)
     SUMA=0

     DIM PESO(13)
     PESO(1)=1
     PESO(2)=2
     PESO(3)=3
     PESO(4)=5
     PESO(5)=7 
     PESO(6)=11
     PESO(7)=13
     PESO(8)=17
     PESO(9)=19
     PESO(10)=21 
     PESO(11)=23
     PESO(12)=29
     PESO(13)=31
     FOR I=1 TO LEN(Cuenta)
       SUMA=SUMA+(CINT(MID(Cuenta,I,1)) * PESO(I))
     NEXT 
     RESTO = TRIM (CSTR(SUMA MOD 11))  
     IF RESTO <> 10 then 
      RESTO = RESTO
     Else 
      RESTO = MID(Cuenta,7,1)
     END IF           
   End Function   
%>
<script Language="JavaScript">
function FrontPage_Form1_Validator(theForm)
{
  if (theForm.Imp.value.length > 16)
  {
    alert("Por favor escriba a lo sumo 16 caracteres en el campo \"Importe\".");
    theForm.Imp.focus();
    return (false);
  }

  var chkVal = theForm.Imp.value;
  var prsVal = parseFloat(theForm.Imp.value);
  if (chkVal != "" && !(prsVal > "0"))
  {
    alert("Por favor escriba un número mayor que \"0\" en el campo \"Importe\".");
    theForm.Imp.focus();
    return (false);
  }
  
  if ( theForm.Imp.value == "" )
  {
    alert("Por favor escriba un valor en el campo \"Importe\".");
    theForm.Imp.focus();
    return (false);
  } 
   
  var prstope = parseFloat( theForm.Tope.value );
  if ( prsVal > prstope )
  {
    alert("No puede poner un importe mayor que el saldo de la cuenta mas los intereses."); 
    return (false);  
  }
  
  var prsInteres = parseFloat( theForm.Interes.value );
  if ( prsVal < prsInteres )
  {
    alert("No puede poner un importe menor que el valor de los intereses."); 
    return (false);  
  }    

  return (true);
}
</script>

<form method="POST" name="Form" action="Amortizar.asp" onsubmit="return FrontPage_Form1_Validator(this)">
  <input type="hidden" name="hname" value="hvalue">
  <input type="hidden" name="Cta_Credito" value="<%=Cta_Credito%>">
  <input type="hidden" name="Interes" value="<%=Interes%>">  
  <input type="hidden" name="Tope" value="<%=Tope%>">

    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="5">
      <tr>
        <td width="47%" bgcolor="#FFCCCC" valign="top" align="right">
          <span class="style3">Cuenta a Amortizar:</span>
        </td>
        <td width="50%" bgcolor="#FFCCCC" valign="top"><span class="style3"><%=Cta_Credito%></span></td>
      </tr>
      <tr>
        <td width="47%" bgcolor="#FFCCCC" valign="top" align="right">
        <span class="style3">Cuenta a Debitar:</span>
        </td>
        <td width="50%" bgcolor="#FFCCCC" valign="top">
          <p align="left"><select name="Cta_Debito" size="1" tabindex="1">
          <% do while not rs.EOF%>
          <% if (mid(rs("cue_subcue"),1,3) <> "141") and (mid(rs("cue_subcue"),1,3) <> "151") then %>
          <%  query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
              set rs2 = conn.Execute( query2 ) 
              if not rs2.eof then
                money = rs2("Cod_Moneda") 
              else
                money = "??"
              end if 
		      cta = money & sucu & mid(rs("cue_subcue"), 3, 1) & whois & rs("des_cuenta") %>
              <option value="<%=cta%><%=Resto(cta)%>"><%=cta%><%=Resto(cta)%></option>
          <% end if %>
          <% rs.MoveNext %><% loop %></select>
          </p>        </td>
      </tr>      
      <tr>
        <td width="47%" height="25" bgcolor="#FFCCCC" align="right">
          <span class="style3">Importe:</span>        </td>
        <td width="50%" height="25" bgcolor="#FFCCCC">
         <span class="style3">$</span> <input type="text" name="Imp" size="18" onkeyup="solo_num(Imp)" Value="<%=FormatNumber(Importe,2,-1,0,0)%>"></td>
      </tr>
      <tr>
        <td width="100%" valign="middle" height="40" colspan="2">
          <p align="center"><input type="submit" value="AMORTIZAR" name="B1" style="color: #FFFFFF; background-color: #9D2C4A; font-family: Verdana; font-weight: bold; border-style: outset"></td>
      </tr>
    </table>
</form>
<% Else
      ' Ejecuta la amortización de préstamo a partir de los datos seleccionados
	  Importe = Request.Form("Imp")
	  Interes = Request.Form("Interes")
	  Cta_Debito = Request.Form("Cta_Debito")
	  Cta_Credito = Request.Form("Cta_Credito")

      tony = "Amortizando: " & Cta_Credito & " Fecha: " & Date() & " Inter: " & Interes*100 & ";"
      ltony = len(tony)
      Impt = Importe * 100
      X_MENSENV="0200"
      X_BBITMAP="F23880018A0180000000000006000000"           
      X_PAN=SPACE(15) & whois 
      X_PRCODE="511100"   
      X_AMOUNTRA=RIGHT("000000000000" & Impt,12) 
      X_TDT=RIGHT("00" & CSTR(month(date)), 2)+RIGHT("00" & CSTR(day(date)), 2)+RIGHT("00" & CSTR(hour(time)), 2)+RIGHT("00" & CSTR(minute(time)), 2)+RIGHT("00" & CSTR(second(time)), 2)
      COUNTER=Application("NumVisits") 
      X_SYSTRACE=RIGHT("000000" & Application("NumVisits"),6)
      X_TIMELOCA=MID(X_TDT,5,6)
      X_DATELOCA=MID(X_TDT,1,4)
      X_DATECAPT=MID(X_TDT,1,4)
      X_ACINCODE="  000000001"
      X_FOINCODE="  000000001"
      X_NUM_UNIQ=CSTR(YEAR(date)) & RIGHT("00" & CSTR(month(date)), 2) & RIGHT("00" & CSTR(day(date)), 2) & X_SYSTRACE
      X_RETRENUM=MID(X_NUM_UNIQ, 3, 12)
      X_RESPCODE="00"
      X_CURRCODE="840"  
          
      ELECTR = HEX (whois) & HEX (X_TDT) & HEX (Request.Form("Imp") * 100) & HEX(hour(Time)) & HEX(minute(Time)) & HEX(second(Time))
      
      'Firmar el comprobante digitalmente

      set AspObj = Server.CreateObject("Comp.CompCrypt")
      F_ELECTR = AspObj.Firmar( ELECTR ) 
	  F_ELECTR = mid(F_ELECTR, 1, 20)
      
      lenfe=len(F_ELECTR) 
      ltony=ltony+len(F_ELECTR)+14 
      X_OBSEV= Right("000" & ltony,3) & tony & "Comprobante:" & Right("00" & lenfe,2) & F_ELECTR
      X_ACCOIDE1="28" & Cta_Debito & "000000001  001" 
      X_ACCOIDE2="28" & Cta_Credito & "000000001  001" 

      query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
      set rs = conn.Execute( query )

      X_DIRECCIO=rs("DIR_TDX25")   
      X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_OBSEV & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2

	  query1 = "INSERT INTO M_MENSWI (ID_MENSAJE, PAN, TDT, SYSTRACE, DES_MENSCE, DIR_TDX25, NUM_UNIQCE, ACINCODE, TIMELOCA, DATELOCA, DATECAPT, RETRENUM, ENVIADO, CODIGO, IDCARCOD, IDCARTER, ACTIONCODE, FILENAME, RESPCODE, BIT48RESP, FEC_ACCION, PROCESADO  ) VALUES ( '" & X_MENSENV & "', '" & X_PAN & "', '" & X_TDT & "', '" & X_SYSTRACE & "', '', '" & X_DIRECCIO & "', '" & X_NUM_UNIQ & "', '" & X_ACINCODE & "', '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_DATECAPT & "', '" & X_RETRENUM & "', '0' ,'11','','','','','','',{//},'')"
      conn.Execute( query1 )
	  
	   ' Ahora lo que hacemos es actualizar ese campo
         Set rsMessage = Server.CreateObject( "ADODB.RecordSet" )
		 
		 rsMessage.CursorType 		= 2 'adOpenKeyset
		 rsMessage.CursorLocation 	= 2 'adUseServer
		 rsMessage.Locktype 		= 3 'adLockDynamic  	
		 rsMessage.ActiveConnection = conn
		 rsMessage.Open "m_menswi"
		 rsMessage.Filter = "SYSTRACE = '" & X_SYSTRACE & "'"	 
         rsMessage("DES_MENSCE")= X_MENSAJE
		 rsMessage.UPDATE 

%>
<IFRAME ALIGN=MIDDLE FRAMEBORDER=0 HEIGHT=400 WIDTH=100% NORESIZE SCROLLING=NO allowtransparency="True" SRC="Amortizar_Procesando.asp?count=37&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>&Cta_debito=<%=Cta_Debito%>&Cta_credito=<%=Cta_Credito%>&importe=<%=Importe%>&Firma=<%=F_ELECTR%>"></iframe>
<%End IF%>
</body>
</html>