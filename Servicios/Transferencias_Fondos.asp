<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<meta http-equiv="Page-Exit" content="revealTrans(Duration=1,Transition=23)">
<title>Transferencias de Fondos</title>
<script language="JavaScript">
function solo_num(theComp)
{
  var checkOK = "0123456789-.";
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
  { //alert("Por favor escriba un numero válido en este campo \"Importe\". El separador decimal permitido es el punto. No necesita separar los miles.");
    theComp.value = allNum
    theComp.focus();
    return (false);
  }
  return (true);
}
</script>
<style type="text/css">
<!-- 
.style1 {
	font-family: Verdana;
	font-size: 12px;
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

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body style="background-color:transparent;">
<script Language="JavaScript" Type="text/javascript">
function Validator(theForm)
{

  if (theForm.Cta_Credito.value == "")
  {
    alert("Por favor, especifíque la Cuenta a Acreditar.");
    theForm.Cta_Credito.focus();
    return (false);
  }

  if (theForm.Cta_Credito.value.length < 14)
  {
    alert("La Cuenta a Acreditar debe tener 14 dígitos.");
    theForm.Cta_Credito.focus();
    return (false);
  }

  if (theForm.Cta_Credito.value.length > 14)
  {
    alert("La Cuenta a Acreditar debe tener 14 dígitos.");
    theForm.Cta_Credito.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.Cta_Credito.value;
  var allValid = true;
  var validGroups = true;
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
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Entre solo dígitos en el campo Cuenta a Creditar.");
    theForm.Cta_Credito.focus();
    return (false);
  }

  if (theForm.Imp.value == "")

  {
    alert("Entre un valor en el Campo Importe.");
    theForm.Imp.focus();
    return (false);
  }
  if (theForm.Imp.value.length > 16)
  {
    alert("No más de 16 Caracteres en el Campo Importe.");
    theForm.Imp.focus();
    return (false);
  }

  var checkOK = "0123456789-.";
  var checkStr = theForm.Imp.value;
  var allValid = true;
  var validGroups = true;
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
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Entre solo dígitos en el campo importe.");
    theForm.Imp.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Entre un número válido en el campo importe.");
    theForm.Imp.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseFloat(allNum);
  if (chkVal != "" && !(prsVal >= "0"))
  {
    alert("Entre un valor mayor que 0 en el campo importe.");
    theForm.Imp.focus();
    return (false);
  }

  if (theForm.Factura.value == "")
  {
    alert("Entre un valor para el campo Factura.");
    theForm.Factura.focus();
    return (false);
  }

  if (theForm.Factura.value.length < 3)
  {
    alert("Entre al menos 3 caracteres (s/n) para el campo factura.");
    theForm.Factura.focus();
    return (false);
  }

  if (theForm.Concepto_Pago.value == "")
  {
    alert("Especifíque el concepto de pago.");
    theForm.Concepto_Pago.focus();
    return (false);
  }

  if (theForm.Concepto_Pago.value.length < 10)
  {
    alert("Entre al menos 10 caracteres en el concepto de pago.");
    theForm.Concepto_Pago.focus();
    return (false);
  }
  if (theForm.Paguese_a.value == "")
  {
    alert("Entre un valor para el campo Paguese A.");
    theForm.Paguese_a.focus();
    return (false);
  }

  if (theForm.Paguese_a.value.length < 5)
  {
    alert("Entre al menos 5 caracteres para el campo Paguese A.");
    theForm.Paguese_a.focus();
    return (false);
  }
  return (true);
}
</script>
<!-- #Include file = "Informa.asp"-->
<%  set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4)

   Cuemay = "CUE" & sucu
   Application.Lock 
   Application("NumVisits") = Application("NumVisits") + 1 
   Application.Unlock 

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

  
  query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM " & Cuemay & " WHERE (Cod_Contra = '" & whois & "') AND (Cue_Inact <> .T.) AND (Cue_Cierre<> .T.) AND (Cue_bloq <> .T.) AND ((CUE_SUBCUE = '3210') OR (CUE_SUBCUE = '3280') OR (CUE_SUBCUE = '3290'))"
  set rs = conn.Execute( query )
%>

<form METHOD="POST" ACTION="Transferencias_Fondos.asp" onsubmit="return Validator(this)" name="FrontPage_Form1">
  <input type="hidden" name="hname" value="hvalue">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style11">Transferencia de Fondos</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
</table>
<table border="0" width="97%" cellspacing="12" cellpadding="5" style="background-color:transparent; ">
<%empresa=rs("Nom_client")%>
    <tr>
      <td align="left" valign="top" height="22" bgcolor="#FFD6D6"><font face="Verdana" color="#000000" size="2"><strong><%=empresa%></strong></font></td>
    </tr>
    <tr>
      <td align="left" valign="top" bgcolor="#FFD6D6"><font color="#000000" size="2" face="verdana" style="font-weight:bold; ">Cuenta a Debitar:</font>
	    &nbsp;&nbsp;&nbsp;&nbsp;
		<select name="Cta_Debito" size="1" tabindex="1">
<% do while not rs.EOF
 
    query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
    'query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = 'CUP')" 
	      set rs2 = conn.Execute( query2 ) 
		  money = rs2("Cod_Moneda") 
		  cta = money & sucu & mid(rs("cue_subcue"), 3, 1) & whois & rs("des_cuenta") %>
         <option value="<%=cta%><%=Resto(cta)%>"><%=cta%><%=Resto(cta)%></option>
<% rs.MoveNext 
   loop %>      
 </select>&nbsp;<small><small><font color="#000000" size="4"></font></small></small></font></td>
    </tr>
    <tr align="center" bgcolor="#FFD6D6">
      <td align="left" valign="top"><font face="Verdana" color="#000000" size="2"><strong>Cuenta a Acreditar: 
	  <input SIZE="15" NAME="Cta_Credito" tabindex="2" MAXLENGTH="14" onkeyup="solo_num(this)">       &nbsp; Importe:
      <input SIZE="15" NAME="Imp" tabindex="3" MAXLENGTH="16" onkeyup="solo_num(Imp)"></strong></font></td>
    </tr>
    <tr align="left" bgcolor="#FFD6D6">
      <td height="30" valign="top"><font face="Verdana" color="#000000" size="2"><strong>Nro. Factura: 
	    <input NAME="Factura" id="Factura" tabindex="4" SIZE="20" MAXLENGTH="20">
      </strong></font></td>
    </tr>
    <tr align="left" bgcolor="#FFD6D6">
      <td height="30" valign="top"><font face="Verdana" color="#000000" size="2"><strong>Concepto de Pago: 
	  <input NAME="Concepto_Pago" id="Concepto_Pago" style="font-family: Verdana; width: 100%" tabindex="5">
      </strong></font></td>
    </tr>
    <tr align="center" bgcolor="#FFD6D6">
      <td align="left" valign="top" bgcolor="#FFD6D6"><font face="Verdana" color="#000000" size="2"><strong>P&aacute;guese a: </strong></font><font face="Verdana" color="#000000" size="2">
        <input name="Paguese_a" type="text" style="font-family: Verdana; width: 100%" tabindex="6" value="">
</font></td>
    </tr>
    </tr>
    <tr align="center">
      <td align="center" valign="top"><input name="submit" type="submit" style="background-color: #9D2C4A; color: rgb(255,255,255); font-weight: bold" tabindex="7" value="Transferir">

<input name="reset" type="reset" style="background-color: #9D2C4A; color: rgb(255,255,255); font-weight: bold" tabindex="8" value="Limpiar"></td>
    </tr>
    <tr align="center">
      <td align="center" valign="top" bgcolor="#FCE8AB"><span class="style1"><strong>Nota:</strong> De no existir el N&uacute;mero de Factura entonces debe especificar <strong>S/N</strong> (Sin N&uacute;mero).</span></td>
    </tr>
  </table>
  </center>
  </div>
</form>
<% Else
   ' Ejecuta la transferencia a partir de los datos seleccionados
   ' Verifica que el cliente este autorizado a acreditar cuentas presupuestadas
    qPresup = "SELECT * FROM PRESUP WHERE ( SUCURSAL = '" & sucu & "' ) AND ( CLIENTE = '" & whois & "' )"
    set rsPresup = conn.Execute( qPresup )    
   NO3210 = mid(Request.Form("Cta_Debito"), 7, 1) <> "1"
  if (rsPresup.EOF or NO3210) and ((mid(Request.Form("Cta_Credito"), 7, 1) = "8") or (mid(Request.Form("Cta_Credito"), 7, 1) = "9")) then %>
<table width="100%" align="center" cellpadding="5" bgcolor="#FCE8AB">
   <tr>
    <td align="left"><span class="style12">Error.  </span>
  </tr>
  <tr>
    <td align="center"><hr size="1">  
  </tr>
  <tr>
    <td><p align="center"><blink><font face="Verdana">
    <font color="#FF0000">
    <b>No es posible acreditar una cuenta presupuestada. Por favor, consulte con el banco.</b>
    </font> </font></blink></p>
    </td>
  </tr>
    <tr>
    <td align="center" class="style14"><a href="Transferencias_Fondos.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
<% else If Request.Form("Cta_Debito").Count = 0 Then %>
<table width="100%" align="center" cellpadding="5" bgcolor="#FCE8AB">
   <tr>
    <td align="left"><span class="style12">Error.  </span>
  </tr>
  <tr>
    <td align="center"><hr size="1">  
  </tr>
  <tr>
    <td><p align="center"><blink><font face="Verdana">
    <font color="#FF0000">
    <b>Usted no ha seleccionado ninguna Cuenta!!!.</b>
    </font> </font></blink></p>
    </td>
  </tr>
    <tr>
    <td align="center" class="style14"><a href="Transferencias_Fondos.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
<% Else 
   cta_cr = mid(Request.Form("Cta_Credito"), 1 ,13)
   rest_cr = resto(cta_cr)
   bank = mid(Request.Form("Cta_Credito"), 3 ,4)
 
   query9 = "SELECT COD_BANCO, SIST_SABIC FROM C_BANCOS WHERE (COD_BANCO = '" & bank & "' ) AND (SIST_SABIC=.T.) AND (TRANSITO=.T.)"
   set rs9 = conn.Execute( query9 ) 
  
   If (mid(Request.Form("Cta_Credito"), 1, 2) = mid(Request.Form("Cta_Debito"), 1, 2) OR ((mid(Request.Form("Cta_Credito"), 1, 2) = "43") and (mid(Request.Form("Cta_Debito"), 1, 2) = "01")) ) AND mid(Request.Form("Cta_Credito"), 14 ,1) = rest_cr AND not rs9.EOF Then 
      ' Cuentas con igual moneda  digito de chequeo ok  banco correcto  
       queryCE = "SELECT ACT_CE FROM M_ESTSIS"
       set rsCE = conn.Execute( queryCE )
	   If rsce("ACT_CE") = True Then  
      
      ' Determinar si se debe cobrar comision o no
      qComi = "SELECT * FROM NOCOMISI WHERE (SUCURSAL='" & sucu & "') AND (CLIENTE='" & whois & "') AND (CUENTA='"& Request.Form("Cta_Debito") &"')"
      set rsComi = conn.Execute( qComi )
   
      if (not rsComi.EOF) or ((not rsPresup.EOF) and ((mid(Request.Form("Cta_Credito"), 7, 1) = "8") or (mid(Request.Form("Cta_Credito"), 7, 1) = "9"))) then
        ' No cobrar la comision
        X_ACINCODE="  000000002"
      else
        ' Cobrar la comision
        X_ACINCODE="  000000001"
      end if
 
      tony = "Pagando Factura: "& Request.Form("Factura")&"; Por Concepto de: "&Request.Form("Concepto_Pago")&"; PAGUESE A: "&Request.Form("Paguese_a")
      ltony = len(tony)
      Impt = Request.Form("Imp") * 100
      X_MENSENV="0200"
      X_BBITMAP="F23880018A0180000000000006000000"           
      X_PAN=SPACE(15) & whois 
      X_PRCODE="510000"   
      X_AMOUNTRA=RIGHT("000000000000" & Impt,12) 
      X_TDT=RIGHT("00" & CSTR(month(date)), 2)+RIGHT("00" & CSTR(day(date)), 2)+RIGHT("00" & CSTR(hour(time)), 2)+RIGHT("00" & CSTR(minute(time)), 2)+RIGHT("00" & CSTR(second(time)), 2)
        COUNTER=Application("NumVisits") 
      X_SYSTRACE=RIGHT("000000" & Counter,6)
      X_TIMELOCA=MID(X_TDT,5,6)
      X_DATELOCA=MID(X_TDT,1,4)
      X_DATECAPT=MID(X_TDT,1,4)    
      X_FOINCODE="  000000001"
      X_NUM_UNIQ=CSTR(YEAR(date)) & RIGHT("00" & CSTR(month(date)), 2) & RIGHT("00" & CSTR(day(date)), 2) & X_SYSTRACE
      X_RETRENUM=MID(X_NUM_UNIQ, 3, 12)
      X_RESPCODE="00"
      X_CURRCODE="840"  
      ELECTR = HEX (whois) & HEX (X_TDT) & HEX (Request.Form("Imp") * 100)
      
      'Firmar el comprobante digitalmente

      set AspObj = Server.CreateObject("Comp.CompCrypt")
      F_ELECTR = AspObj.Firmar( ELECTR )
	  F_ELECTR = mid(F_ELECTR, 1, 20)

      lenfe=len(F_ELECTR) 
      ltony=ltony+len(F_ELECTR)+14 
      X_OBSEV= Right("000" & ltony,3) & tony & "Comprobante:" & Right("00" & lenfe,2) & F_ELECTR
      X_ACCOIDE1="28" & Request.Form("Cta_Debito")  & "000000001  001" 
      X_ACCOIDE2="28" & Request.Form("Cta_Credito") & "000000001  001" 

      query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
      set rs = conn.Execute( query )

      X_DIRECCIO=rs("DIR_TDX25")
	  RS.Close   
      X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_OBSEV & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2
     	 
	  queryInsert = "INSERT INTO M_MENSWI (ID_MENSAJE, PAN, TDT, IDCARCOD, IDCARTER, SYSTRACE, DES_MENSCE, Actioncode, Filename, Respcode, bit48resp, DIR_TDX25, NUM_UNIQCE, ACINCODE, TIMELOCA, DATELOCA, DATECAPT, RETRENUM, ENVIADO, CODIGO, FEC_ACCION, Procesado) VALUES ( '" & X_MENSENV & "', '" & X_PAN & "', '" & X_TDT & "', '', '', '" & X_SYSTRACE & "', '' , '', '', '', '', '" & X_DIRECCIO & "', '" & X_NUM_UNIQ & "', '" & X_ACINCODE & "', '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_DATECAPT & "', '" & X_RETRENUM & "', '0', '51', {//}, '')"
	  conn.Execute(queryInsert) 
	
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
<iframe frameborder="0" height="100%" width="100%" allowtransparency="true" scrolling="auto" src="Transferencias_Fondos_Proces.asp?count=37&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>&Cta_debito=<%=Request.Form("Cta_Debito")%>&Cta_credito=<%=Request.Form("Cta_Credito")%>&importe=<%=Request.Form("Imp")%>&ELECTR=<%=F_ELECTR%>"></iframe>
<%Else%>
<table width="100%" cellpadding="5" cellspacing="5" bgcolor="#FCE8AB">
  <tr>
    <td align="left"><span class="style12">Error.  </span>
  </tr>
  <tr>
    <td align="center"><hr size="1">  
  </tr>
  <tr>
    <td align="center"><span class="style13">No existe
    conexión con el banco!!!</span></td>
  </tr>`
    <tr>
    <td align="center" class="style14"><a href="Transferencias_Fondos.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>	
    <%End If
	 Else %> 
<table width="100%" cellpadding="5" cellspacing="5" bgcolor="#FCE8AB">
  <tr>
    <td align="left"><span class="style12">Error.  </span>  </tr>
  <tr>
    <td align="center"><hr size="1">  </tr>
  <tr>
    <td align="center"><span class="style13">Cuenta a acreditar incorrecta o las monedas no coinciden.  
  </span></tr>
  <tr>
    <td align="center" class="style14"><a href="Transferencias_Fondos.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
<% End If  End If  End If  End if %> </td>

</body>
</html>