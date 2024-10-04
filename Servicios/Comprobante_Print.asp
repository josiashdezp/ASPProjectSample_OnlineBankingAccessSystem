<% Response.Expires = 0
   Response.Buffer = True %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Recuperación de Comprobantes de Transferencia</title>
 <style>
 
 body { scrollbar-arrow-color: #DDCE67;
scrollbar-base-color: #f0b6c4;
scrollbar-face-color: #9f314e;
scrollbar-highlight-color: #f0b6c4;
scrollbar-shadow-color: #DDCE67; }

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
	color: #000000;
	font-size: 18px;
	filter:glow(color=#FFFFFF,strength=1);
width:100%;
}
.style5 {
	font-family: Verdana;
	color: #000000;
	font-size: 10px;
	
}
.style6 {
	font-family: Verdana;
	color: #000000;
	font-size: 16px;
	
}
 </style>
<style>
		A		{
			text-decoration : none;
			color : "#CC0000";
			font-size : 16px;		}

		A:hover		{
			text-decoration : underline;
					}
	</style>
</head>
<body style="background-color: transparent" leftmargin="0">
<!-- #Include file = "Informa.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style1">Recuperar Comprobante.</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
  
</table>
<% 
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4)

   set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")  

   query1 = "SELECT DISTINCT FECHA FROM M_TRAVB WHERE (substr(CTA_DB, 3, 4)= '" & sucu & "') AND (substr(CTA_DB, 8, 4)= '" & whois & "') AND ((RESULTADO = 'Transacción contabilizada satisfactoriamente!!!') OR (REF_CORRIE='00000000') ) Order By Fecha DESC"  
   set rs1 = conn.Execute( query1 ) 
   
   'On Error Resume Next
   If (Request.Form("hname") = "") AND (Request.Form("hname1")) = "" Then %>

<form method="POST" action="Comprobante.asp">
  <input type="hidden" name="hname" value="hvalue">
    <table width="100%%"  border="0" cellspacing="2" cellpadding="2">
      <tr>
        <td width="40%" align="left"><span class="style2"><strong>Fecha a recuperar:</strong></span></td>
        <td width="30%" align="left"><select size="1" name="date">
<% do while not rs1.EOF%>
     <option value="<%=rs1("fecha")%>"><%=rs1("fecha")%></option>
<% rs1.MoveNext
     loop %>
  </select></td>
        <td width="30%" align="left"><input type="submit" value="Buscar" name="B1" style="color: #FFFFFF; font-family: Verdana; font-weight: bold; background-color: #9D2C4A; border-style: outset"></td>
      </tr>
    </table>
</form>
<%  Response.Write Request.Form("hname1")
 ELSE
 If Request.Form("hname1") = "" Then 
 date1 = Request.Form("date") 
 date1 = year(date1) & "/" & month(date1) & "/" & day(date1)  
 queryf = "SELECT * FROM M_TRAVB WHERE (FECHA={^" & date1 & "}) AND (substr(CTA_DB, 3, 4)= '" & sucu & "') AND (substr(CTA_DB, 8, 4)= '" & whois & "') AND ((RESULTADO = 'Transacción contabilizada satisfactoriamente!!!') OR (REF_CORRIE='00000000') )"
 set rsf = conn.Execute( queryf )  %>
<form method="POST" action="Comprobante.asp">
  <input type="hidden" name="hname1" value="hvalue">
    <span class="style2"><strong>Seleccione un Comprobante:</strong></span><br><br>
  <select name="Ray" size="1">
<% Do while not rsf.EOF%>    
<option value="<%tony = "Importe: " & Csng(rsf("IMPORTE")) & " Referencia:  " & rsf("REF_CORRIE") & " Firma: " & rsf("FIRMA") %><%=tony%>"><%=tony%></option>
<%rsf.MoveNext
   Loop%>
    </select> <br><br>
	<input type="submit" value="Recuperar" name="B1" style="font-family: Verdana; color: #FFFFFF; font-weight: bold; background-color: #9D2C4A; border-style: outset">
<% Else
  cadena = Request.Form("Ray") 
   pos = instr(cadena, "Referencia:") + 13 
 pos1 = instr(cadena, "Importe:") 
 pos2 = (pos-14) - (pos1+9) 
 pos3 = instr(cadena, "Firma:")-1 
 firma = Trim( mid( cadena, pos3+7 ) ) 

  Set regEx = New RegExp            
  regEx.Pattern = ","
  regEx.IgnoreCase = True          
  Importe = regEx.Replace(MID(CADENA, 10, POS2),".")

  queryf = "SELECT * FROM M_TRAVB WHERE ( substr(CTA_DB, 3, 4)= '" & sucu & "') AND (substr(CTA_DB, 8, 4)= '" & whois & "') AND (REF_CORRIE	 = '" & mid(cadena, pos, pos3-pos) & "') AND (IMPORTE =" & Importe & ") AND (FIRMA = '"& Firma &"')"
  'Response.Redirect("Prueba.asp?Data="&queryf&"&vall="&vall)
  set rsf = conn.Execute(queryf)
  
  select case rsf("Codigo") 
  case 51 'Comprobante de Transferencia %>
<table border="1" width="100%" cellpadding="0" bordercolor="#000000" cellspacing="0" align="center">
    <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
              de Crédito y Comercio</font></td>
            <td width="63%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
              de Transferencia.<br>
              </strong></font><font color="#000000" face="Verdana" size="2">Sistema
              de Conexión Cliente-Banco, Virtual Bandec.</font></td>
          </tr>
        </table>
      </td>
    </tr>
    
    <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") or (rsf("RESP_CODE") = "02") Then %>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Fecha Contable:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("fec_contab")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Referencia Corriente:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><%=rsf("ref_corrie")%></font></td>
    </tr>
    <%end if%>
    
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Debitada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_db")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Acreditada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_cr") %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Importe
      Transferido:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;$
        </strong> <% =FormatNumber(rsf("Importe"),2) %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Observaciones:</b></font></td>
      <%      Obs = rsf("obs")        %>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =Obs %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Resultado:</b></font></td>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =rsf("Resultado") %></font>
      </td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Comprobante:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% = rsf("firma") %></font></td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="52">&nbsp;</td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="18"><blockquote>
        <font color="#000000" face="Verdana"><strong><div align="left"><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="2">
          Hecho:
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;
        Autorizado:</font></p>
        </div></strong></font>
      </blockquote>
      </td>
    </tr>
    
  <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") or (rsf("RESP_CODE") = "02") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
  <%End If%>
    <%If rsf("RESP_CODE") <> "00" Then %>
     <%If rsf("RESP_CODE") = "01" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta transferencia no será tramitada hasta el siguiente día hábil, 
         debido a que la sucursal de destino no está conectada a la Red Pública de
         Transmisión de Datos.</font></font></strong>
         </td>
       </tr>
     <%Else%>
     <%If rsf("RESP_CODE") = "02" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta transferencia será tramitada por Correo Electrónico con la sucursal de destino.</font></font></strong>
         </td>
       </tr>
     <%else%>       
       <tr align="center">
       <td width="90%" align="center" colspan="9" height="24" valign="top">
       <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000"><br>
       Esta operación no se ejecutó correctamente.<br>
       Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></font></strong>
       </td>
       </tr> 
     <%End If
     End If
  End If%>  
  </table>
<% case 10 'Comprobante de Compra de Combustible 
	CADENA = rsf("obs")
    
    p = instr( CADENA, "CFin:" )
    p1 = instr( p, CADENA, ";" )
    CFIN = mid( CADENA, p+5, p1-p-5 )

    p = instr( CADENA, "Cant_Tarj:" )
    p1 = instr( p, CADENA, ";" )
    Cant_Tarj = mid( CADENA, p+10, p1-p-10 )
	Impf = rsf("Importe") + (cint(Cant_Tarj)*8)
%>
<div align="left"><table border="1" width="100%" cellpadding="0" bordercolor="#000000" cellspacing="0" align="center">
    <tr>
      <td width="100%" align="middle" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="10%"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="25%"><font face="Verdana" size="2">Banco de Crédito y
              Comercio</font></td>
            <td width="65%" align="center"><font face="Verdana"><b>Comprobante
              de Compra de Combustible</b><br>
              <font size="2">Sistema de Conexión Cliente-Banco, Virtual Bandec.</font></font></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="60%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Fecha Contable:</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("fec_contab")%></font></td>
    </tr>
    <tr>
      <td width="60%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Referencia Corriente:</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("ref_corrie")%></font></td>
    </tr>
    <tr>
      <td width="60%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cliente de Fincimex:</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =cfin%></font></td>
    </tr>
    <tr>
      <td width="60%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta Debitada:</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><%=rsf("Cta_Db") %></font></td>
    </tr>
    <tr>
      <td width="60%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Importe para Compra de Combustible + Importe por Compra de Tarjetas (FINCIMEX):</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =FormatNumber(Impf,2) %></font></td>
    </tr>
    <tr>
      <td width="60%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Aporte a la Reserva Estatal (INRE):</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =FormatNumber(cdbl(Impf/100),2)%></font></td>
    </tr>
    <tr>
	  <td width="60%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Cantidad de Tarjetas Compradas:</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =Cant_Tarj %>  Por un Importe de <strong>&nbsp;$</strong> <%=FormatNumber(cint(Cant_Tarj)*8,2)%> </font></td>
    </tr>
    <tr>
      <td width="60%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Resultado
      de la Operación:</b></font></td>
      <td width="40%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =rsf("Resultado")%></font>
      </td>
    </tr>
    <tr>
      <td width="60%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Comprobante de la Operación:</b></font></td>
      <td width="40%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("Firma")%></font></td>
    </tr>
    <tr>
      <td width="100%" align="middle" colspan="2" height="52">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" align="middle" colspan="2" height="18"><blockquote>
        <font color="#000000" face="Verdana"><strong><div align="left"><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          Hecho:
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        Autorizado:</p>
        </div></strong></font>
      </blockquote>
      </td>
    </tr>
    <%If rsf("RESP_CODE") <> "00" Then 
        If rsf("RESP_CODE") = "01" Then %>
      <tr>
      <td width="90%" align="center" colspan="2" height="46"><strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br></small>Esta operación no será tramitada hasta el siguiente día hábil, <small><br>
      </small>debido a que la sucursal de destino no está conectada a la Red Pública de Transmisión de Datos.</font></strong></td>
      </tr>
     <%Else%>
      <tr>
      <td width="90%" align="center" colspan="2" height="46"><strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br></small>Esta operación no se ejecutó correctamente.<small><br>
      </small>Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></strong></td>
      </tr> 
     <%End If
     Else%>
     <tr>
     <td width="90%" align="center" colspan="2" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
    </tr>
   <%End If%>
  </table>
  </div>
<% case 58 'Comprobante de los aportes al presupuesto 
    CADENA = rsf("obs")
    p = instr( CADENA, "NIT:" )
    p1 = instr( p, CADENA, ";" )
    NIT = mid( CADENA, p+4, p1-p-4 )
    p = instr( CADENA, "SUC:" )   
    if p <> 0 then
      p1 = instr( p, CADENA, ";" )
      SUC = mid( CADENA, p+4, p1-p-4 )
    else
      SUC = sucu
    end if    
  
    p = instr( CADENA, "PD:" )
    p1 = instr( p, CADENA, ";" )
    PD= mid( CADENA, p+3, p1-p-3 )
  
    p = instr( CADENA, "PH:" )
    p1 = instr( p, CADENA, ";" )
    PH= mid( CADENA, p+3, p1-p-3 )
    
    p = instr( CADENA, "TP:" )
    p1 = instr( p, CADENA, ";" )
    TP= mid( CADENA, p+3, p1-p-3 )
    
    p = instr( CADENA, "RF:" )
    p1 = instr( p, CADENA, ";" )
    RF= mid( CADENA, p+3, p1-p-3 )       
        
    p = instr( CADENA, "II:" )
    p1 = instr( p, CADENA, ";" )
    II= FormatNumber( mid( CADENA, p+3, p1-p-3 )/100 )
                   
    p = instr( CADENA, "Principal:" )
    p1 = instr( p, CADENA, ";" )
    pcpal= FormatNumber( mid( CADENA, p+10, p1-p-10 )/100 )
    
    p = instr( CADENA, "Recargo:" )
    p1 = instr( p, CADENA, ";" )
    Recargo= FormatNumber( mid( CADENA, p+9, p1-p-9 )/100 )      
    
    p = instr( CADENA, "Multa:" )
    p1 = instr( p, CADENA, ";" ) 
    Multa= FormatNumber( mid( CADENA, p+7, p1-p-7 )/100 )
    
    p = instr( CADENA, "TI:" )
    p1 = instr( p, CADENA, ";" )
    TI= FormatNumber( mid( CADENA, p+3, p1-p-3 )/100 )
    
    p = instr( CADENA, "IO:" )
    p1 = instr( p, CADENA, ";" )
    IO= FormatNumber( mid( CADENA, p+3, p1-p-3 )/100 )
        
    p = instr( CADENA, "PF:" )
    p1 = instr( p, CADENA, ";" )
    PF= mid( CADENA, p+3, p1-p-3 )
    
    Parrafo = mid(PF, 8)
    PF = mid(PF, 1,7)
%> 
<div align="center"><center>
<table border="1" width="100%" cellpadding="0" bordercolor="#000000" bordercolordark="#000000" cellspacing="0">
  <tr>
    <td width="90%" align="center" colspan="9" height="39">
      <table border="0" cellpadding="5" cellspacing="5" width="100%">
        <tr>
          <td width="8%"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
          <td width="23%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
            de Crédito y Comercio</font></td>
          <td width="69%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
            de Aporte.<br>
            </strong></font><font color="#000000" face="Verdana" size="2">Sistema
            de Conexión Cliente-Banco, Virtual Bandec.</font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="9" height="44"><p align="left"><font color="#000000" face="Verdana"><strong>ONAT&nbsp;&nbsp;&nbsp;
      <font size="1">Declaración&nbsp;
    </font></strong></font><input type="checkbox" name="C1" value="ON">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <font face="Verdana" size="1">&nbsp;&nbsp; Número Identificación Tributaria</font><font color="#000000" face="Verdana" size="1"><strong>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CR
    </strong></font><%If mid(Parrafo,7,1)=2 then%> <small><font face="Verdana"><strong>- 03</strong></font></small> <%else%> <font face="Verdana"><small><strong>- 04</strong></small></font> <%End IF%><font color="#000000" face="Verdana" size="1"><strong><br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Pago&nbsp; </strong></font><input type="checkbox" name="C1" value="ON" checked>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;<font face="Verdana"><%=NIT%></font></td>
  </tr>
  <tr><% Cuemay = "CUE" & sucu %>
      <td width="58%" align="right" height="31" colspan="2"><p align="left"><font color="#000000" face="Verdana" size="1">Debítese a:<br>
<%query = "SELECT Nom_Client FROM "& Cuemay &" WHERE (Cod_Contra = '" & whois & "')"
set rs = conn.Execute( query ) %>    </font><font color="#000000" face="Verdana" size="2"><%=rs("Nom_Client")%></font><font color="#000000" face="Verdana" size="1"> </font></td>
    <td width="32%" align="center" height="31" colspan="7"><font color="#000000" face="Verdana"><font size="1">Código de la Cuenta:<br>
    </font><% =Rsf("Cta_Db")%></font></td>
  </tr>
  <tr>
    <td width="58%" align="right" height="31" colspan="2"><p align="left"><font color="#000000" face="Verdana" size="1">Tributo:<br>
    </font><font color="#000000" face="Verdana" size="2"><%=Parrafo%></font><font color="#000000" face="Verdana" size="1"> </font></td>
    <td width="32%" align="center" height="31" colspan="7"><font color="#000000" face="Verdana" size="1">Código:<br>
    </font><font face="Verdana"><% =PF %></font></td>
  </tr>
  <tr>
    <td width="58%" align="right" height="38" colspan="2"><p align="left"><font color="#000000" face="Verdana" size="1">Importe:<br>
    </font></td>
    <td width="32%" align="center" height="38" colspan="7">
      <p align="center"><b>$</b>&nbsp; <font color="#000000" face="Verdana"><%=FormatNumber(rsf("Importe"),2)%></font></p>
  </td>
  </tr>
  <tr align="center">
    <td width="39%" align="left" height="169" rowspan="6" valign="top" bordercolor="#000000" bordercolorlight="#000000"><p align="left"><font face="Verdana" size="1">Breve
    explicación de la Referencia de Pago:<br>
    Representa la Clasificación de la declaración y/o pago que se realiza y se Identifica
    mediante:<br>
    0 - Pago Voluntario<br>
    1 - Número de Convenio<br>
    2 - Número de la Resolución de Auditoria<br>
    3 - Declaración Jurada rectificada<br>
    Este código se reflejará en el escaque señalado por las siglas TP (Tipo de Pago) y a
    continuación del mismo se consignará el número de aprobación que ampara dicho pago. Se
    exceptúa el código 0.</font></td>
    <td width="19%" align="center" height="24" valign="top">
      <p align="left"><font face="Verdana" size="1">Principal:</font></p>
  </td>
    <td width="32%" align="center" height="24" valign="top" colspan="7"><font face="Verdana">&nbsp;<% =Pcpal %></font></td>
  </tr>
  <tr align="center">
    <td width="19%" align="left" height="24" valign="top">
      <p align="left"><font face="Verdana" size="1">Recargo:</font></p>
    </td>
    <td width="32%" align="left" height="24" valign="top" colspan="7">
      <p align="center"><font face="Verdana">&nbsp;<% =Recargo %></font></td>
  </tr>
  <tr align="center">
    <td width="19%" align="left" height="24" valign="top"><font face="Verdana" size="1">Multa
    o Sanción:</font></td>
    <td width="32%" align="left" height="24" valign="top" colspan="7">
      <p align="center"><font face="Verdana">&nbsp;<% =Multa %></font></p>
 </td>
  </tr>
  <tr align="center">
    <td width="19%" align="left" height="50" valign="top" rowspan="2"><font face="Verdana" size="1">Período a Liquidar:</font></td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><font face="Verdana" size="1">D</font> </td>
    <td width="5%" align="center" height="24" valign="top" bordercolor="#000000"><font face="Verdana" size="1">M</font> </td>
    <td width="5%" align="center" height="24" valign="top" bordercolor="#000000"><font face="Verdana" size="1">A</font> </td>
    <td width="6%" align="center" height="24" valign="top" bordercolor="#000000">&nbsp; </td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><font face="Verdana" size="1">D</font> </td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><font face="Verdana" size="1">M</font> </td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><font face="Verdana" size="1">A</font> </td>
  </tr>
  <tr align="center">
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><%=mid(PD,1,2)%>
</td>
    <td width="5%" align="center" height="24" valign="top" bordercolor="#000000"><%=mid(PD,4,2)%>
</td>
    <td width="5%" align="center" height="24" valign="top" bordercolor="#000000"><%=mid(PD,7,4)%>
</td>
    <td width="6%" align="center" height="24" valign="top" bordercolor="#000000">&nbsp; </td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><%=mid(PH,1,2)%>
</td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><%=mid(PH,4,2)%>
</td>
    <td width="4%" align="center" height="24" valign="top" bordercolor="#000000"><%=mid(PH,7,4)%>
</td>
  </tr>
  <tr align="center">
    <td width="32%" align="left" height="39" valign="top"><font face="Verdana" size="1">Referencia
    de Pago:</font></td>
    <td colspan="2" valign="top" align="left" height="39"><p align="center"><font face="Verdana" size="1">TP<br>
<%=TP%>    </font></td>
    <td colspan="5" valign="top" align="left" height="39"><p align="center"><font face="Verdana" size="1">D<br>
    </font><font face="Verdana">
<%=RF%>    </font></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="9" height="49" valign="top">
      <table border="1" width="100%" height="50" align="right"  cellspacing="0" bordercolorlight="#000000" cellpadding="0">
      <tr>
        <td width="42%" align="center"><font face="Verdana" size="1">Importe de la Base Imponible</font></td>
        <td width="10%" align="center"> <font face="Verdana" size="1">TI</font></td>
        <td width="48%" align="center"><font face="Verdana" size="1">Importe de la Obligación</font></td>
      </tr>
      <tr>
        <td width="42%" ><p align="center"><font face="Verdana" size="2"><%=II%></font><font face="Verdana" size="1"> </font></td>
        <td width="10%"><p align="center"><font face="Verdana" size="2"><%=TI%></font><font face="Verdana" size="1"> </font></td>
        <td width="48%"><p align="center"><font face="Verdana" size="1">&nbsp;<b>$</b> </font><font face="Verdana" size="2"><%=IO%></font><font face="Verdana" size="1"> </font></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Resultado:&nbsp;</small><font size="2"><b><% = rsf("Resultado") %></b></font></font></td>
  </tr>
  <%If instr( CADENA, "SUC:" )<>0  and (sucu <> SUC) Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Tramitado por correo electrónico en la sucursal:&nbsp;</small><font size="2"><b><%=SUC%></b></font></font></td>
  </tr>
  <%End if
  If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") or (rsf("RESP_CODE") = "02") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Referencia:&nbsp;</small><font size="2"><b><% = rsf("ref_corrie") %></b></font></font></td>
  </tr>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Fecha contable:&nbsp;</small><font size="2"><b><% =rsf("fec_contab") %></b></font></font></td>
  </tr>
  <%end if%>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><p align="center"><font face="Verdana"><small>Comprobante:
      </small></font><font color="#000000" face="Verdana"><b><% = rsf("Firma") %></b></font></td>
  </tr>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="37"><blockquote>
      <font color="#000000" face="Verdana"><strong><p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Hecho:
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Autorizado:</font></p>
      </strong></font>
    </blockquote>
    </td>
  </tr>
  <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") or (rsf("RESP_CODE") = "02") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
  <%End If
  If rsf("RESP_CODE") <> "00" Then 
     If rsf("RESP_CODE") = "01" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta operación no será tramitada hasta el siguiente día hábil, 
         debido a que la sucursal de destino no está conectada a la Red Pública de
         Transmisión de Datos.</font></font></strong>
         </td>
       </tr>
     <%Else%>
     <%If rsf("RESP_CODE") = "02" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta operación será tramitada por Correo Electrónico con la sucursal de destino.</font></font></strong>
         </td>
       </tr>
     <%Else%>     
       <tr align="center">
       <td width="90%" align="center" colspan="9" height="24" valign="top">
       <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
       Esta operación no se ejecutó correctamente.<br>
       Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></font></strong>
       </td>
       </tr> 
     <%End If
     End If     
  End If%>  
 </table>
</center></div>  
<% case 11 'Comprobante de amortizacion de prestamos 
 CADENA = rsf("obs") 

 p = instr( CADENA, "Inter:" ) 
   if p <> 0 then
     p1 = instr(p, CADENA, ";")
     if p1 <> 0 then
       Interes = mid(CADENA, p+7, p1-p-7)/100
     else
       Interes = 0
     end if
   else
     Interes = 0
   end if    
   
   Importe = rsf("Importe")   
   
   if Importe >= Interes then
     Principal = Importe - Interes
   else
     Interes = Importe
     Principal = 0
   end if
%>
<table border="1" width="100%" cellpadding="0" bordercolor="#000000" cellspacing="0" align="center">
    <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
              de Crédito y Comercio</font></td>
            <td width="63%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
              de Amortización.<br>
              </strong></font><font color="#000000" face="Verdana" size="2">Sistema
              de Conexión Cliente-Banco, Virtual Bandec.</font></td>
          </tr>
        </table>
      </td>
    </tr>
    <%If (rsf("RESP_CODE") = "00") Then %>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Fecha Contable:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("fec_contab")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Referencia Corriente:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><%=rsf("ref_corrie")%></font></td>
    </tr>
    <%end if%>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Debitada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_db")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Acreditada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_cr") %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Principal amortizado:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;$
        </strong> <% =FormatNumber(Principal,2) %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Interés amortizado:</b></font></td>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong>&nbsp;$&nbsp;</strong><% =FormatNumber(Interes,2) %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Resultado:</b></font></td>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =rsf("Resultado") %></font>
      </td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Comprobante
      de la Transferencia:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% = rsf("firma") %></font></td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="52">&nbsp;</td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="18"><blockquote>
        <font color="#000000" face="Verdana"><strong><div align="left"><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="2">
          Hecho:
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;
        Autorizado:</font></p>
        </div></strong></font>
      </blockquote>
      </td>
    </tr>
  <%If (rsf("RESP_CODE") = "00") Then %>
      <tr align="center">
        <td width="90%" align="center" colspan="9" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
      </tr>
  <%Else%> 
      <tr align="center">
      <td width="90%" align="center" colspan="9" height="24" valign="top">
      <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000"><br>
      Esta operación no se ejecutó correctamente.<br>
      Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></font></strong>
      </td>
      </tr> 
  <%End If%>
  </table>

<% case 12 'Comprobante de acreditacion de nomina 
 case 08 'Comprobante de solicitud de chequera
   query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rsf("resp_code") & "')"
   set rs3 = conn.Execute( query3 )    
  %>
  <table border="1" width="90%" cellpadding="0" bordercolor="#000000" bordercolordark="#000000" cellspacing="0" ID="Table1">
  <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%" ID="Table2">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
              de Crédito y Comercio</font></td>
            <td width="63%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
              de Solicitud de Chequera.<br>
              </strong></font><font color="#000000" face="Verdana" size="2">Sistema
              de Conexión Cliente-Banco, Virtual Bandec.</font></td>
          </tr>
        </table>
      </td>
    </tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Fecha
    de Solicitud:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("FECHA")%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Cuenta:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("Cta_Db")%></strong></font></td>
  </tr>
  <%TipCheq = RIGHT(rsf("CTA_CR"), 3)
  
    If TipCheq = "001" then
       TipCheq = "Nominativo"
    Else
    If TipCheq = "101" then
       TipCheq = "Nominativo a la Orden"    
    Else
    If TipCheq = "003" then
       TipCheq = "Certificado"
    Else
    If TipCheq = "103" then
       TipCheq = "Certificado a la Orden"
    Else  
    If TipCheq = "010" then  
       TipCheq = "Vaucher Nominativo"
    End If
    End If
    End If
    End If
    End If        
  %>
  
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Tipo de Chequera Solicitada:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<% = TipCheq%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Cantidad Solicitada:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong><%=FormatNumber(rsf("IMPORTE"),0)%></strong></font></td>
  </tr>  
  <tr>
    <td width="40%" align="right" height="19"><font color="#000000" face="Verdana">Resultado
    de la Solicitud:</font></td>
    <% If rsf("resp_code") = "00" then
         resultado = "Solicitud aprobada, Chequera en confección"
       Else
         resultado = rs3("spanish")
       End If
    %>
    <td width="50%" align="left" height="19"><font color="#000000" face="Verdana"><strong>&nbsp;<%=resultado%></strong></font>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Comprobante
    de la Solicitud:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("FIRMA")%></strong></font></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="52"></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="18"><blockquote>
      <font color="#000000" face="Verdana"><strong><p align="left">Hecho:
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Autorizado:</p>
      </strong></font>
    </blockquote>
    </td>
  </tr>
  <tr>
  </table>
<% case 01 'Remesas de efectivo (SEPSA) %>
<table border="1" width="100%" cellpadding="0" bordercolor="#000000" cellspacing="0" align="center">
    <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
              de Crédito y Comercio</font></td>
            <td width="63%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
              de Remesa de efectivo.<br>
              </strong></font><font color="#000000" face="Verdana" size="2">Sistema
              de Conexión Cliente-Banco, Virtual Bandec.</font></td>
          </tr>
        </table>
      </td>
    </tr>
    
    <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") Then %>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Fecha Contable:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("fec_contab")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Referencia Corriente:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><%=rsf("ref_corrie")%></font></td>
    </tr>
    <%end if%>
    
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Acreditada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_cr") %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Sucursal
      Destino:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cue_sucur")%></font></td>
    </tr>    
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Importe
      Transferido:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;$
        </strong> <% =FormatNumber(rsf("Importe"),2) %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Observaciones:</b></font></td>
      <%
      Obs = rsf("obs")
      Obs = mid( Obs, instr( Obs, "PAGUESE A:" ) )
      
      if instr( Obs, "Comprobante:" ) > 0 then
        Obs = mid( Obs, 1, instr(Obs, "Comprobante:")-1 )
      end if
            
      %>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =Obs %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Resultado
      de la operación:</b></font></td>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =rsf("Resultado") %></font>
      </td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Comprobante
      de la operación:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% = rsf("firma") %></font></td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="52">&nbsp;</td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="18"><blockquote>
        <font color="#000000" face="Verdana"><strong><div align="left"><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="2">
          Hecho:
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;
        Autorizado:</font></p>
        </div></strong></font>
      </blockquote>
      </td>
    </tr>
    
  <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") or (rsf("RESP_CODE") = "02") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
  <%End If
  If rsf("RESP_CODE") <> "00" Then 
     If rsf("RESP_CODE") = "01" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta operación no será tramitada hasta el siguiente día hábil, 
         debido a que la sucursal de destino no está conectada a la Red Pública de
         Transmisión de Datos.</font></font></strong>
         </td>
       </tr>
     <%Else
     If rsf("RESP_CODE") = "02" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta operación será tramitada por Correo Electrónico con la sucursal de destino.</font></font></strong>
         </td>
       </tr>
     <%Else%>     
       <tr align="center">
       <td width="90%" align="center" colspan="9" height="24" valign="top">
       <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
	   Esta operación no se ejecutó correctamente.<br>
       Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></font></strong>
       </td>
       </tr> 
     <%End If%>
     <%End If%>     
  <%End If%>  
  </table>
  </div> 
<% case 03 'Depositos de efectivo (SEPSA) %>
<table border="1" width="100%" cellpadding="0" bordercolor="#000000" cellspacing="0" align="center">
    <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
              de Crédito y Comercio</font></td>
            <td width="63%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
              de Depósito de efectivo.<br>
              </strong></font><font color="#000000" face="Verdana" size="2">Sistema
              de Conexión Cliente-Banco, Virtual Bandec.</font></td>
          </tr>
        </table>
      </td>
    </tr>
    <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") Then %>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Fecha Contable:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("fec_contab")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Referencia Corriente:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><%=rsf("ref_corrie")%></font></td>
    </tr>
    <%end if%>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Debitada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_db")%></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Cuenta
      Acreditada:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% =rsf("cta_cr") %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Importe
      Transferido:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;$
        </strong> <% =FormatNumber(rsf("Importe"),2) %></font></td>
    </tr>
    <%
      Obs = rsf("obs")
      pos = instr( Obs, "CIRMON:" )
      
      if pos <> 0 then 
      
        Cirmon = mid( Obs, pos+7, 3 )
        
        queryCM = "SELECT NOM_CIRMON FROM C_CIRMON WHERE COD_CIRMON='" & Cirmon & "'"
        set rsCM = conn.Execute( queryCM )
        
        Cirmon = Cirmon & " - " & rsCM( "NOM_CIRMON" )
        Obs = mid( Obs, pos+10 )
    %>
    <tr>
      <td width="40%" align="right" height="18">
      <font color="#000000" face="Verdana" size="2"><b>Código de Circulación Monetaria:</b></font>
      </td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2">
      &nbsp;<% =Cirmon %></font>
      </td>
    </tr>      
    <%     
      end if
           
      if instr( Obs, "Comprobante:" ) > 0 then
        Obs = mid( Obs, 1, instr(Obs, "Comprobante:")-1 )
      end if
    %>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Observaciones:</b></font></td>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =Obs %></font></td>
    </tr>
    <tr>
      <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2"><b>Resultado
      de la operación:</b></font></td>
      <td width="50%" align="left" height="19"><font face="Verdana" size="2"><strong><font color="#000000">&nbsp;</font></strong><% =rsf("Resultado") %></font>
      </td>
    </tr>
    <tr>
      <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2"><b>Comprobante
      de la operación:</b></font></td>
      <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;</strong><% = rsf("firma") %></font></td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="52">&nbsp;</td>
    </tr>
    <tr>
      <td width="90%" align="center" colspan="2" height="18"><blockquote>
        <font color="#000000" face="Verdana"><strong><div align="left"><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="2">
          Hecho:
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;
        Autorizado:</font></p>
        </div></strong></font>
      </blockquote>
      </td>
    </tr>
  <%If (rsf("RESP_CODE") = "00") or (rsf("RESP_CODE") = "01") or (rsf("RESP_CODE") = "02") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
  <%End If
  If rsf("RESP_CODE") <> "00" Then 
     If rsf("RESP_CODE") = "01" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta operación no será tramitada hasta el siguiente día hábil, 
         debido a que la sucursal de destino no está conectada a la Red Pública de
         Transmisión de Datos.</font></font></strong>
         </td>
       </tr>
     <%Else
     If rsf("RESP_CODE") = "02" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
         Esta operación será tramitada por Correo Electrónico con la sucursal de destino.</font></font></strong>
         </td>
       </tr>
     <%Else%>     
       <tr align="center">
       <td width="90%" align="center" colspan="9" height="24" valign="top">
       <strong><font face="Verdana" size="2"><font color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><br>
       Esta operación no se ejecutó correctamente.<br>
       Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></font></strong>
       </td>
       </tr> 
     <%End If
     End If     
  End If%>  
  </table>
  </div>
<% case 13 'Orden de cobro %>   
<table border="1" width="100%" cellpadding="0" bordercolor="#000000" bordercolordark="#000000" cellspacing="0">
  <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="bottom"><font color="#000000" face="Verdana" size="2">Banco
              de Crédito y Comercio</font></td>
            <td width="63%" align="center"><font color="#000000" face="Verdana" size="3"><strong>Comprobante
              de Orden de Cobro.<br>
              </strong></font><font color="#000000" face="Verdana" size="2">Sistema
              de Conexión Cliente-Banco, Virtual Bandec.</font></td>
          </tr>
        </table>
      </td>
    </tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Fecha
    Contable:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("fec_contab")%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Referencia
    Corriente:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("ref_corrie")%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Cuenta
    Acreditada:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<% =rsf("cta_db") %></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Importe
    Total:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>$ <% =FormatNumber(rsf("Importe"),2) %></strong></font></td>
  </tr>
  <tr>  
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Observaciones:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("Obs")%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="19"><font color="#000000" face="Verdana">Resultado
    de la Operación:</font></td>
    <td width="50%" align="left" height="19"><font color="#000000" face="Verdana"><strong>&nbsp;<% =rsf("resultado") %></strong></font>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Firma del comprobante:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=rsf("firma")%></strong></font></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="52"></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="18"><blockquote>
      <font color="#000000" face="Verdana"><strong><p align="left">Hecho:
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Autorizado:</p>
      </strong></font>
    </blockquote>
    </td>
  </tr>
  <tr>
<%If rsf("RESP_CODE") <> "00" Then %>

<%If rsf("RESP_CODE") = "01" Then %>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="46"><strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br>
    </small>Esta transferencia no será tramitada hasta el siguiente día hábil, <small><br>
    </small>debido a que la sucursal de destino no está conectada a la Red Pública de
    Transmisión de Datos.</font></strong></td>
  </tr>
<%Else%>
  <tr>
    <td width="90%" align="center" colspan="2" height="46"><strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br>
    </small>Esta operación no se ejecutó correctamente.<small><br>
    </small>Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></strong></td>
  </tr> 
<%End If
 Else%>
  <tr>
    <td width="90%" align="center" colspan="2" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
<%End If%>
    </table>
  </div>
<%
 end select 
End If
END IF%> 
</form>
</body>
</html>