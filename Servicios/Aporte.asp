<%  Response.Expires=0 
    Response.Buffer = True  
   %>  <html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>Aporte al Presupuesto</title>
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

function parrafo_moneda( FormComp )
{
  var cta = new String( FormComp.Cta_Debito.value );
 
  if ( cta.substring(0,2) == '40' ) 
  {
    FormComp.DIV.style.visibility="hidden"
    FormComp.DIV.style.display="none"
    FormComp.MN.style.visibility="visible"
    FormComp.MN.style.display="inline"
    FormComp.Parrafo.value = FormComp.MN.value
  }
  else 
  {
    FormComp.MN.style.visibility="hidden"
    FormComp.MN.style.display="none"
    FormComp.DIV.style.visibility="visible"
    FormComp.DIV.style.display="inline"    
    FormComp.Parrafo.value = FormComp.DIV.value
  };
    
  return (true);
}

function parrafo_cambio( FormComp )
{
  var cta = new String( FormComp.Cta_Debito.value );
 
  if ( cta.substring(0,2) == '40' ) 
  {
    FormComp.Parrafo.value = FormComp.MN.value
  }
  else 
  {
    FormComp.Parrafo.value = FormComp.DIV.value
  };
    
  return (true);
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script>

</head>
<body leftmargin="0" <%if Request.Form("hname") = "" then%>onload="parrafo_moneda(FrontPage_Form1)" <%end if%> style="background-color:transparent;">
<% 
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4) 

   'position=inStr(Request.ServerVariables("Logon_User"),"\") 
   'sucu = mid(Request.ServerVariables("Logon_User"),position+1,4)
   'whois = mid(Request.ServerVariables("Logon_User"),position+5,4)
      
   Cuemay = "CUE" & sucu
   Application.Lock 
   Application("NumVisits") = Application("NumVisits") + 1 
   Application.Unlock 

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
   
   Function Formalize( date_part )
     If CInt( date_part ) < 10 then
       Formalize = "0" & date_part
     Else
       Formalize = date_part
     End If
   End Function
  
  yy = DatePart( "yyyy", Date() )
  'On Error Resume Next
   If Request.Form("hname") = "" Then
          'Permite la captación de datos en la forma HTML
   
  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
  
  query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM "& Cuemay &" WHERE (Cod_Contra = '" & whois & "') AND ((CUE_SUBCUE = '3210') OR (CUE_SUBCUE = '3280') OR (CUE_SUBCUE = '3290'))"
  set rs = conn.Execute( query )
   
  queryp = "SELECT Sig_Moneda, Cod_Ingpre, Parrafo FROM c_ingpre ORDER BY Cod_Ingpre"
  set rsp = conn.Execute( queryp ) 
 
  querym = "SELECT COD_BANCO, NOM_BANCO1 FROM C_BANCOS WHERE (ELECTRON = .T.) ORDER BY COD_PROVIN, COD_BANCO"
  set rsm = conn.Execute( querym ) 
  Nombre=rs("Nom_Client")%>
<script Language="JavaScript">
<!--
function FrontPage_Form1_Validator(theForm)
{ /********************************************************************
		/* la primera validación es para el campo NIT */
 
  if (theForm.NIT.value == "")
  {
    alert("Please enter a value for the \"Numero de Identificación Tributaria\" field.");
    theForm.NIT.focus();
    return (false);
  }

  if (theForm.NIT.value.length < 11)
  {
    alert("Please enter at least 11 characters in the \"Numero de Identificación Tributaria\" field.");
    theForm.NIT.focus();
    return (false);
  }

  if (theForm.NIT.value.length > 11)
  {
    alert("Please enter at most 11 characters in the \"Numero de Identificación Tributaria\" field.");
    theForm.NIT.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.NIT.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  
  /***  Este for para revisar la cadena solo números. */
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))         break;
	  
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  
  
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Numero de Identificación Tributaria\" field.");
    theForm.NIT.focus();
    return (false);
  }
  
  /*   fin para el campo NIT */
  /*******************************************************************************/
  /*  ahora se valida el campo listbox de las cuentas a debitar */
 
  if (theForm.Cta_Debito.selectedIndex < 0)
  {
    alert("Please select one of the \"Cuenta  a Debitar\" options.");
    theForm.Cta_Debito.focus();
    return (false);
  }

/*/** esto es para el importe *****************************************************/
     

  if (theForm.Imp.value == "")
  {
    alert("Please enter a value for the \"Importe\" field.");
    theForm.Imp.focus();
    return (false);
  }

  if (theForm.Imp.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Importe\" field.");
    theForm.Imp.focus();
    return (false);
  }

  if (theForm.Imp.value.length > 16)
  {
    alert("Please enter at most 16 characters in the \"Importe\" field.");
    theForm.Imp.focus();
    return (false);
  }
 
 /********************************************************************************/
 /* validación del campo lista llamado MN */
 
  if (theForm.MN.selectedIndex < 0)
  {
    alert("Please select one of the \"Parrafo\" options.");
    theForm.MN.focus();
    return (false);
  }


 /********************************************************************************/
 /* esta es la del campo Principal */

  var checkOK = "0123456789-.";
  var checkStr = theForm.Pcpal.value;
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
      allNum += ".";
      decPoints++;
    }
    else
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Pcpal\" field.");
    theForm.Pcpal.focus();
    return (false);
  }

    if (theForm.Pcpal.value == "")
  {
    alert("Please enter a value for the \"Principal\" field.");
    theForm.NIT.focus();
    return (false);
  }
  
  if (decPoints > 1)
  {
    alert("Please enter a valid number in the \"Pcpal\" field.");
    theForm.Pcpal.focus();
    return (false);
  }
  
  /******** Recargo y Multa ***********************/
  
  
 if ((theForm.Rcgo.value == "") ||(theForm.Multa.value == ""))
  {
    if(confirm("Uno de los campos \"Recargo\" o \"Multa\" está vacío. Desea que Bandec Online le asigne valor 0."))
	{ 
	theForm.Rcgo.value='0';
	theForm.Multa.value='0';
	}
	else
	{
	if (theForm.Rcgo.value == "")
		 theForm.Rcgo.focus();
	else
		theForm.Multa.focus();
	return (false);
	}
  }

 
 
  /******** este es el de importe Base Imponible ***********************/
  
   if (theForm.Imp_Impo.value == "")
  {
    alert("Please enter a value for the \"Importe Base Imponible\" field.");
    theForm.Imp_Impo.focus();
    return (false);
  }

  if (theForm.Imp_Impo.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Importe Base Imponible\" field.");
    theForm.Imp_Impo.focus();
    return (false);
  }

  if (theForm.Imp_Impo.value.length > 16)
  {
    alert("Please enter at most 16 characters in the \"Importe Base Imponible\" field.");
    theForm.Imp_Impo.focus();
    return (false);
  }


 /******** TI % ***********************/
  
   if (theForm.TI.value == "")
  {
    alert("Please enter a value for the \"TI\" field.");
    theForm.TI.focus();
    return (false);
  }

  if (theForm.TI.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"TI\" field.");
    theForm.TI.focus();
    return (false);
  }

  if (theForm.TI.value.length > 16)
  {
    alert("Please enter at most 16 characters in the \"TI\" field.");
    theForm.TI.focus();
    return (false);
  }



/**** esta es para importe de la obligación ***/

  if (theForm.Imp_Obl.value == "")
  {
    alert("Please enter a value for the \"Importe de la Obligación\" field.");
    theForm.Imp_Obl.focus();
    return (false);
  }

  if (theForm.Imp_Obl.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Importe de la Obligación\" field.");
    theForm.Imp_Obl.focus();
    return (false);
  }

  if (theForm.Imp_Obl.value.length > 16)
  {
    alert("Please enter at most 16 characters in the \"Importe de la Obligación\" field.");
    theForm.Imp_Obl.focus();
    return (false);
  }
  

  
  return (true);
}


//--></script>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana;
	font-weight: bold;
	color: #FFFFFF;
	font-size: 20px; 
	filter:glow(color=#000000,strength=2);
    width:100%;
}
.style2 {
	font-family: Verdana;
	color: #FFFFCC;
	font-size: 18px;
	filter:glow(color=#000000,strength=2);
width:100%;
} -->
</style>

<form METHOD="POST" ACTION="Aporte.asp" onSubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
  <input type="hidden" name="hname" value="hvalue"><div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style11">Aportes al Presupuesto</span></td>
  </tr>
    <tr>
    <td width="100%"><hr size="1" color="#FCE8AB"></td>
  </tr>
 </table>
  <table border="0" width="100%" cellspacing="8" cellpadding="7" align="center">
    <tr bgcolor="#FFD6D6">
      <td colspan="3" align="left" valign="middle" ><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><b><font face="Verdana" color="#000000" size="2">NIT:</font></b></td>
          <td><font face="Verdana" color="#000000" size="2"><strong>Deb&iacute;tese a: </strong></font>
            <!--webbot bot="Validation" S-Display-Name="Cuenta  a Debitar" B-Value-Required="TRUE" -->
            <b><font face="Verdana" size="2" color="#000000"> </font></b> <strong><font face="Verdana" color="#000000" size="2">&nbsp;</font></strong></td>
          <td><strong><font face="Verdana" color="#000000" size="2">Importe: </font></strong></td>
        </tr>
        <tr>
          <td><b><font face="Verdana" color="#000000" size="2">
            <input name="NIT" type="text" tabindex="1" size="12" maxlength="11" >
          </font></b></td>
          <td><b><font face="Verdana" size="2" color="#000000">
            <select name="Cta_Debito" size="1" tabindex="2" onChange="parrafo_moneda(FrontPage_Form1)">
              <% do while not rs.EOF        
  query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
    set rs2 = conn.Execute( query2 ) 
    money = rs2("Cod_Moneda") 
	cta = money & sucu & mid(rs("cue_subcue"), 3, 1) & whois & rs("des_cuenta") %>
              <option value="<%=cta%><%=Resto(cta)%>"><%=cta%><%=Resto(cta)%></option>
              <% rs.MoveNext 
 loop %>
            </select>
          </font></b></td>
          <td><b><font face="Verdana" size="2" color="#000000"><strong><font face="Verdana" color="#000000" size="2">$ </font></strong>
                <input name="Imp" tabindex="3" onKeyUp="solo_num(this)" size="12" maxlength="16">
          </font></b></td>
        </tr>
      </table></td>
      </tr>
    <tr>
      <td height="25" colspan="3" align="center" valign="top"><span class="style2" style="font-size:small;font-weight:bold;"><% =Nombre%></span></td>
    </tr>
    <tr bgcolor="#FFD6D6">
      <td height="30" colspan="3" align="left" valign="middle" ><div align="left">
        <font face="Verdana" color="#000000" size="2"><strong>Sucursal en que se va a aportar: <strong></font>
        <select name="SUC" tabindex="4">
		<%Do while not rsm.EOF%>
		<option 
		<%if rsm("Cod_banco")=sucu then%>selected<%end if%> value="<%=rsm("Cod_BANCO")%>"><%=rsm("Cod_BANCO")%>&nbsp;&nbsp;<%=rsm("NOM_BANCO1")%>
		</option>
		<%rsm.MoveNext%><% loop %>
		</select>		</td>  
    </tr>
    <tr bgcolor="#FFD6D6">
      <%rs.MoveFirst%>
      <td height="30" colspan="3" align="left" valign="middle" ><div align="left">        <p><strong>
      <font face="Verdana" color="#000000" size="2">&nbsp;Concepto: </font></strong><!--webbot bot="Validation" S-Display-Name="Parrafo" B-Value-Required="TRUE" -->
      <select name="MN" tabindex="5" onChange="parrafo_cambio(FrontPage_Form1)" <%if rs("sig_moneda") <> "CUP" then%>style="visibility:hidden;display:none" <%end if%>><% Do while not rsp.EOF%>
      <%if mid(rsp("Cod_Ingpre"),7,1) = "2" then %><option value="<%=rsp("Cod_Ingpre")%>&nbsp;&nbsp;<%=rsp("Parrafo")%>"><%=rsp("Cod_Ingpre")%>&nbsp;&nbsp;<%=rsp("Parrafo")%></option><%end if%><% rsp.MoveNext %><% loop %></select><select name="DIV" size="1" tabindex="5" onChange="parrafo_cambio(FrontPage_Form1)" <%if rs("sig_moneda") = "CUP" then%>style="visibility:hidden;display:none" <%end if%>><% rsp.MoveFirst %>
      <% Do while not rsp.EOF 
	  if mid(rsp("Cod_Ingpre"),7,1) = "1" then %><option value="<%=rsp("Cod_Ingpre")%>&nbsp;&nbsp;<%=rsp("Parrafo")%>"><%=rsp("Cod_Ingpre")%>&nbsp;&nbsp;<%=rsp("Parrafo")%></option><%end if%><% rsp.MoveNext %><% loop %></select></td>
    </tr>    
    <tr align="center" bgcolor="#FFD6D6">
      <td colspan="3" align="left" valign="middle" bgcolor="#FFD6D6" ><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><font face="Verdana" color="#000000" size="2"><b>&nbsp;Principal:</b></font></td>
          <td><b><font face="Verdana" color="#000000" size="2">Recargo:</font></b></td>
          <td><font face="Verdana" color="#000000" size="2"><b>Multa:</b></font></td>
        </tr>
        <tr>
          <td><font face="Verdana" color="#000000" size="2"><b>$</b></font>
            <!--webbot bot="Validation" S-Data-Type="Number" S-Number-Separators="x." -->
            <input name="Pcpal" type="text" tabindex="6" onKeyUp="solo_num(this)" size="12">
&nbsp;</td>
          <td><b><font face="Verdana" color="#000000" size="2">$</font></b>
            <!--webbot bot="Validation" S-Data-Type="Number" S-Number-Separators="x." -->
            <input name="Rcgo" type="text" tabindex="7" onKeyUp="solo_num(this)" value="0" size="12"></td>
          <td><font face="Verdana" color="#000000" size="2"><b>$</b></font>
            <!--webbot bot="Validation" S-Data-Type="Number" S-Number-Separators="x." -->
            <input name="Multa" type="text" tabindex="8" onKeyUp="solo_num(this)" value="0" size="12"></td>
        </tr>
      </table>        </td>
    </tr>
    <tr align="center" bgcolor="#FFD6D6">
      <td colspan="3" align="center" valign="middle" >
	  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="4"><b><font face="Verdana" color="#000000" size="2">Período a Liquidar:</font></b><b><font face="Verdana" color="#000000" size="2"> </font></b> </td>
          </tr>
        <tr>
          <td><b><font face="Verdana" color="#000000" size="2">Desde:</font></b></td>
          <td><select name="Day_d" size="1" tabindex="9">
            <% for d = 1 to 31 %>
            <% if d = CInt(dd) then %>
            <option selected value="<%=Formalize(d)%>"><%=d%></option>
            <% else %>
            <option value="<%=Formalize(d)%>"><%=d%></option>
            <% end if %>
            <% next %>
          </select>
/
<select name="Month_d" size="1" tabindex="10">
  <% for m = 1 to 12 %>
  <% if m = CInt(mm) then %>
  <option selected value="<%=Formalize(m)%>"><%=m%></option>
  <% else %>
  <option value="<%=Formalize(m)%>"><%=m%></option>
  <% end if %>
  <% next %>
</select>
/
<select name="Year_d" size="1" tabindex="11">
  <% for y = 1999 to Year(Date()) %>
  <% if y = CInt(yy) then %>
  <option selected value="<%=y%>"><%=y%></option>
  <% else %>
  <option value="<%=y%>"><%=y%></option>
  <% end if %>
  <% next %>
</select>
&nbsp;</td>
          <td><font face="Verdana" color="#000000" size="2"><b>Hasta:</b></font> </td>
          <td><select name="Day_h" size="1" tabindex="12">
            <% for d = 1 to 31 %>
            <% if d = CInt(dd) then %>
            <option selected value="<%=Formalize(d)%>"><%=d%></option>
            <% else %>
            <option value="<%=Formalize(d)%>"><%=d%></option>
            <% end if %>
            <% next %>
          </select>
/
<select name="Month_h" size="1" tabindex="13">
  <% for m = 1 to 12 %>
  <% if m = CInt(mm) then %>
  <option selected value="<%=Formalize(m)%>"><%=m%></option>
  <% Else %>
  <option value="<%=Formalize(m)%>"><%=m%></option>
  <% end if %>
  <% next %>
</select>
/
<select name="Year_h" size="1" tabindex="14">
  <% for y = 1999 to Year(Date()) %>
  <% if y = CInt(yy) then %>
  <option selected value="<%=y%>"><%=y%></option>
  <% else %>
  <option value="<%=y%>"><%=y%></option>
  <% end if %>
  <% next %>
</select></td>
        </tr>
      </table></td>
    </tr>
    <tr align="center" bgcolor="#FFD6D6">
      <td height="30" colspan="3" align="left" valign="middle" ><b><font face="Verdana" color="#000000" size="2">&nbsp;TP (Tipo de Pago)&nbsp; </font></b>&nbsp;<b><font face="Verdana" color="#000000" size="2">Referencia&nbsp;de Pago: </font></b>
        <select size="1" name="TP" tabindex="15">
        <option value="0" selected>0</option>
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
      </select><input type="text" name="TP_Desc" size="12" tabindex="16" onKeyUp="solo_num(this)"></td>
    </tr>
    <tr align="center" bgcolor="#FFD6D6">
      <td colspan="3" align="left" valign="middle" bgcolor="#FFD6D6" ><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="31%"><b><font face="Verdana" color="#000000" size="2">Imp. Base Imponible:</font></b> </td>
          <td width="29%">&nbsp; <font face="Verdana" color="#000000" size="2"><b>TI: %</b></font></td>
          <td width="40%"><b><font face="Verdana" color="#000000" size="2">Imp. Obligación: $</font></b>
            <!--webbot bot="Validation" S-Display-Name="Importe de la Obligación" S-Data-Type="Number" S-Number-Separators="x." B-Value-Required="TRUE" I-Minimum-Length="1" I-Maximum-Length="16" S-Validation-Constraint="Greater than or equal to" S-Validation-Value="1" --></td>
        </tr>
        <tr>
          <td><input name="Imp_Impo" type="text" tabindex="17" onKeyUp="solo_num(this)" size="12"></td>
          <td><input name="TI" type="text" tabindex="18" onKeyUp="solo_num(this)" size="5"></td>
          <td><input name="Imp_Obl" type="text" tabindex="19" onKeyUp="solo_num(this)" size="12" maxlength="19"></td>
        </tr>
      </table></td>
    </tr>
    <tr align="center">
      <td height="25" colspan="3" align="center" valign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input TYPE="submit" VALUE="Aportar" tabindex="20" style="background-color: #9D2C4A; color: #FFFFFF; font-family: Verdana; font-weight: bold; border-style: outset">        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input TYPE="reset" VALUE="Limpiar" tabindex="21" style="background-color: #9D2C4A; color: #FFFFFF; font-family: Verdana; font-weight: bold; border-style: outset">      </td>
      </tr>
    <tr align="center">
      <td height="25" colspan="3" align="left" valign="top" bgcolor="#FCE8AB"><font face="Verdana" size="2" color="#000000"><b>&nbsp;Breve explicación de la Referencia de
      Pago:</b></font> <p><font size="1" face="Verdana"><b>&nbsp;Representa la Clasificación de
      la declaración y/o pago que se realiza y se identifica mediante:</b></font></p>
      <p><font size="1" face="Verdana"><b>&nbsp;0 - Pago Voluntario<br>
      &nbsp;1 - Número de Convenio<br>
      &nbsp;2 - Número de la Resolución de Auditoria<br>
      &nbsp;3 - Declaración Jurada rectificada</b></font></p>
      <p><font size="1" face="Verdana"><b>&nbsp;Este código se reflejará en el escaque
      señalado por las siglas TP (Tipo de Pago) y a continuación&nbsp;<br>
      &nbsp;del mismo se consignará el número de aprobación que ampara dicho pago. Se
      exceptúa el código 0.</b></font></td>
    </tr>
  </table>
  <%rsp.movefirst
  If rs("sig_moneda") = "CUP" then
     money = "2"
   Else
     money = "1"
   End if
  found = false %>
  <input type="hidden" name="Parrafo" value="<% Do while not rsp.EOF%><%if (mid(rsp("Cod_Ingpre"),7,1) = money) and not found then %><%=rsp("Cod_Ingpre")%>&nbsp;&nbsp;<%=rsp("Parrafo")%><% found = true %><%end if%><% rsp.MoveNext %><% loop %>">
  </center></div>
</form>
<% Else ' Ejecuta la transferencia a partir de los datos seleccionados  %>

<% If Request.Form("Cta_Debito").Count = 0 Then %>
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
    <td align="center" class="style14"><a href="aporte.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
<% Else 
     set conn = Server.CreateObject( "ADODB.Connection" )
     conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")  
     queryp = "SELECT Sig_Moneda, Cod_Ingpre, Parrafo, Cod_otrcon FROM c_ingpre WHERE (Cod_Ingpre='" & mid(Request.Form("Parrafo"),1,7) & "')"
     set rsp = conn.Execute( queryp ) 
  
   If rsp("Sig_Moneda") = "CUP" Then
        mda="40" 
     Else
        mda="43"
     End If 
  
   If mda = mid(Request.Form("Cta_Debito"), 1, 2) Then
      ' Cuentas con igual moneda   
      set conn = Server.CreateObject( "ADODB.Connection" )
        conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
        queryCE = "SELECT ACT_CE FROM M_ESTSIS"
        set rsCE = conn.Execute( queryCE )
    
     If rsce("ACT_CE") = True Then 
        set conn = Server.CreateObject( "ADODB.Connection" )
          conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
 
          If mid( Request.Form( "Cta_Debito" ), 1, 2) = "40" Then
             Parrafo = Request.Form("MN")
          Else
             Parrafo = Request.Form("DIV")
          End IF  
 
          Impt = Request.Form("Imp") * 100
          Impi = Request.Form("Imp_Impo") * 100
          Impo = Request.Form("Imp_Obl") * 100
          Impp = Request.Form("Pcpal") * 100
          Impr = Request.Form("Rcgo") * 100
          Impm = Request.Form("Multa") * 100
          TIt= Request.Form("TI") * 100
          tony = "NIT:" & Request.Form("NIT") & ";PD:" & Request.Form("Day_d") & "/" & Request.Form("Month_d") & "/" & Request.Form("Year_d") & ";PH:" & Request.Form("Day_h") & "/" & Request.Form("Month_h") & "/" & Request.Form("Year_h") & ";TP:" & Request.Form("TP") & ";RF:" & RIGHT("                    " & Request.Form("TP_DESC"),20) & ";II:" & RIGHT("000000000000" & Impi,12) & ";Principal:" & RIGHT("000000000000" & Impp,12) & ";Recargo:" & RIGHT("000000000000" & Impr,12) & ";Multa:" & RIGHT("000000000000" & Impm,12) & ";TI:" & RIGHT("00000" & TIt,5) & ";IO:" & RIGHT("000000000000" & Impo,12) & ";PF:" & Parrafo & ";" & "SUC:" & mid(Request.Form("SUC"),1,4) & ";"
        
          ltony = len(tony)
          X_MENSENV="0200" 
          X_BBITMAP="F23880018A0180000000000006000000"           
          X_PAN=SPACE(15) & whois 
          X_PRCODE="580000"   
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


          ELECTR = HEX (whois) & HEX (X_TDT) & HEX (Request.Form("Imp") * 100)
          
          'Firmar el comprobante digitalmente

          set AspObj = Server.CreateObject("Comp.CompCrypt")
          F_ELECTR = AspObj.Firmar( ELECTR )          
          
          lenfe=len(F_ELECTR) 
          ltony=ltony+len(F_ELECTR)+14 

          X_OBSEV= Right("000" & ltony,3) & tony & "Comprobante:" & Right("00" & lenfe,2) & F_ELECTR
          X_ACCOIDE1="28" & Request.Form("Cta_Debito") & "000000001  001" 
   

          Cta_Cr = rsp("Sig_Moneda") & "38104" & rsp("Cod_Otrcon") & "00"
          X_ACCOIDE2="28" & Cta_Cr & "000000001  001" 

          query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
          set rs = conn.Execute( query )

          X_DIRECCIO=rs("DIR_TDX25")   
          X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_OBSEV & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2
   
          
		  query1 = "INSERT INTO M_MENSWI (ID_MENSAJE, PAN, TDT, IDCARCOD, IDCARTER, SYSTRACE, DES_MENSCE, ACTIONCODE, FILENAME, RESPCODE, BIT48RESP, DIR_TDX25, NUM_UNIQCE, ACINCODE, TIMELOCA, DATELOCA, DATECAPT, RETRENUM, ENVIADO, CODIGO, FEC_ACCION, PROCESADO) VALUES ( '" & X_MENSENV & "', '" & X_PAN & "', '" & X_TDT & "', '', '', '" & X_SYSTRACE & "', '', '', '', '', '', '" & X_DIRECCIO & "', '" & X_NUM_UNIQ & "', '" & X_ACINCODE & "', '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_DATECAPT & "', '" & X_RETRENUM & "', '0', '58', {//}, '')"
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

<IFRAME ALIGN=MIDDLE FRAMEBORDER=0 HEIGHT=700 WIDTH=100% allowtransparency="True" NORESIZE SCROLLING=NO SRC="Aporte_Procesando.asp?count=37&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>&Cta_debito=<%=Request.Form("Cta_Debito")%>&Importe=<%=Request.Form("Imp")%>"></iframe>
<% Else ' CE %>
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
    <td align="center" class="style14"><a href="aporte.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>	
<% End If ' CE 
 Else ' Moneda %>
<table width="100%" cellpadding="5" cellspacing="5" bgcolor="#FCE8AB">
  <tr>
    <td align="left"><span class="style12">Error.  </span>  </tr>
  <tr>
    <td align="center"><hr size="1">  </tr>
  <tr>
    <td align="center"><span class="style13">Cuenta a acreditar incorrecta o las monedas no coinciden.  
  </span></tr>
  <tr>
    <td align="center" class="style14"><a href="aporte.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
<% End If ' Moneda 
 End If 
 End If %>
</body>
</html>