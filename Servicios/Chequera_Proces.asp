<html>
 <head>
<title>Solicitud de Chequera</title>
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
	color: #FF0000;
	font-weight: bold;
	font-size: 20px;
	filter:glow(color=#000000,strength=2);
width:100%;
}
.style3 {
	font-family: Verdana;
	color: #000000;
	font-weight: bold;
	font-size: 16px;
	
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
			font-family:Verdana;
			font-weight:bold;
			color : "#FFFFFF";
			font-size : 14px;
			filter:glow(color=#000000,strength=2);
            width:100%;		}

		A:hover		{
			text-decoration : underline;
			color : "#FCE8AB";		}
	</style>
<script language="vbscript">
sub window_onbeforeprint
B1.style.display = "none"
end sub
sub window_onafterprint
B1.style.display = ""
end sub
</script>
</head>
<body style="background-color: transparent" top="0">
<% 
   the_count = CInt(Request("count")) 
   SYSTRACE = Request("Systrace")
   TDT = Request("TDT")
   
   Cta_Debito = Request("Cta_Debito")
   Cnt = Request("cnt")
   TipCheq = Request("TipCheq")
   
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4)
   
   On Error Resume Next
%>

<%   Set Conn = Server.CreateObject("ADODB.Connection")
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")
   query2 = "SELECT bit48resp, respcode FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (respcode <> ' ') AND (Tdt = '" & TDT & "')"
   set rs2 = conn.Execute( query2 ) 
%>

<% If rs2.EOF Then
      If the_count = 0 then %>
 
</table> 
<table width="90%" border="0" align="center" cellpadding="5" cellSpacing="2">
  <tr>
    <td align="left" vAlign="top" bgcolor="#FCE8AB"><span class="style2">Error de comunicaci&oacute;n con Sucursal <% = sucu %>. </span></td>
  </tr>
  <tr>
      <td width="80%" align="left" vAlign="top" bgcolor="#FCE8AB"><span class="style3">No 
        se puede realizar la transacci&oacute;n en estos momentos ya que no se 
        ha podido establecer comunicaci&oacute;n con la sucursal a la que pertenecen 
        sus cuentas. <br><br> Por favor, pruebe 
        de nuevo o intente m&aacute;s tarde.</span></td>
  </tr>
</table>
	<%else%>
          <meta http-equiv="REFRESH" content="5; url=Chequera_Proces.asp?count=<%= the_count-1%>&Systrace=<%=SYSTRACE%>&TDT=<%=TDT%>&Cta_debito=<%=Cta_Debito%>&Cnt=<%=Cnt%>&TipCheq=<%=TipCheq%>">
    <table width="70%" cellpadding="5" cellspacing="3" align="center">
    <tr bgcolor="#9C2E4B"> 
      <td colspan="3" align="center"><b><font face="Arial" color="white"> 
        <marquee scrollamount="5" width="100%" behavior="alternate">
        <font size="2" face="Verdana">Chequeando Conexi&oacute;n con Sucursal 
        <% = Sucu %>
        ...</font> 
        </marquee>
      </font></b></td>
    </tr>

    <tr> 
       <td align="center"><img src="../images/computer.gif" width="100" height="70"></td>
       <td align="center"><img src="../images/bytes.gif" width="200" height="50"></td>
       <td align="center"><img src="../images/computer.gif" width="100" height="70"></td>
    </tr>
  </table>   
  
  <%End If%>       	     	     	     	     	     	     	     	     
        
   <%Else%>
   
    <html>
  
    <head>
     <title>Solicitud de Chequera</title>
    </head>
  
    <body>

<% Cadena = UCASE(rs2("bit48resp")) %>
<% CADENA1 = rs2("bit48resp") %>

<% 
   query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rs2("respcode") & "')"
   set rs3 = conn.Execute( query3 )    
%>

<div align="center"><center>

<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bordercolordark="#000000">
  <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="5" width="100%">
          <tr>
            <td width="7%" align="left"><img border="0" src="../Images/Logo_Vino.gif" WIDTH="50" HEIGHT="49"></td>
            <td width="30%" valign="top"><font color="#000000" face="Verdana" size="2">Banco
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
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<% = Date%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana">Cuenta:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<% =Cta_Debito %></strong></font></td>
  </tr>
 
  <%codcheq = TipCheq
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
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong><% =FormatNumber(Cnt,0) %></strong></font></td>
  </tr>
  

<%
   
   fin = INSTR(CADENA,"COMPROBANTE:")
   p = (Cint(mid(CADENA, fin+12,2)) - 14)   
   ELECTR = mid( CADENA, fin+14, p )
%>
  
  <tr>
    <td width="40%" align="right" height="19"><font color="#000000" face="Verdana">Resultado
    de la Solicitud:</font></td>
    <% If rs2("respcode") = "00" then
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
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana"><strong>&nbsp;<%=ELECTR%></strong></font></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="40">&nbsp;</td>
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
<%  
   query1 = "INSERT INTO M_TRAVB (FECHA, CTA_DB, CTA_CR, IMPORTE, RESULTADO, RESP_CODE, OBS, FIRMA, TDT, CUE_SUCUR, COD_CONTRA, SIG_MONEDA, CODIGO, REF_CORRIE, FEC_CONTAB) VALUES ( '"& date &"', '"& Cta_Debito & "', '"& RIGHT("00000000000000" & CodCheq, 14) &"', '"& Cnt &"', '"& rs3("spanish") &"', '"& rs2("respcode") &"', '', '"& ELECTR &"', '"& CSTR(YEAR(date)) & RIGHT("00" & CSTR(month(date)), 2) & RIGHT("00" & CSTR(day(date)), 2) &"', '"& sucu &"', '"& whois &"', '"& CUP &"', '08', '00000000', '"& date &"')"
   set rs1 = conn.Execute( query1 )
%>
  
<%If rs2("RESPCODE") <> "00" Then %>

<%If rs2("RESPCODE") = "01" Then %>
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
<%End If%>
<% Else%>
  <tr>
    <td width="90%" align="center" colspan="2" height="18"><font color="#000000" face="Verdana"><strong>Presente este comprobante en su Sucursal para recoger la Chequera solicitada.</strong></font></td>
  </tr>
  </table>
  <Div id="B1">
    <p align="center"><A href="#" onClick="window.open 'Comprobante_Small.asp?Tipo=CHEQ&Cta_Debito=<%=Cta_Debito%>&TipCheq=<%=TipCheq%>&Cnt=<%=Cnt%>&Resultado=<%=resultado%>&ELECTR=<%=ELECTR%>','SubMenu','height=300,width=400,resizable,scrollbars,statusbar'"><< IMPRIMIR COMPROBANTE >>
         </a></P>
</div>
<%End If%>
<%End If%>
</body>
</html>