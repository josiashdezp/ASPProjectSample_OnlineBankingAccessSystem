        <html>
        <head>
		<meta http-equiv="Page-Exit" content="revealTrans(Duration=1,Transition=23)">
        <title>Transferencia de Fondos</title>
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
        <body style="background-color:transparent;">
<% 
   the_count = CInt(Request("count")) 
   SYSTRACE = Request("Systrace")
   TDT = Request("TDT")
   
   Cta_Debito = Request("Cta_Debito")
   Cta_Credito = Request("Cta_Credito")
   Importe = Request("Importe")
   ELECTR = Request("ELECTR")
   
   on error resume next
   set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")

   query2 = "SELECT bit48resp, respcode FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (respcode <> ' ') AND (Tdt = '" & TDT & "')"
   set rs2 = conn.Execute( query2 ) 
   If rs2.EOF Then
   
      If the_count = 0 then 
   'Borra el mensaje 200 del m_menswi {16/04/2001}
 
   query = "DELETE FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (Tdt = '" & TDT & "')"
   set rs2 = conn.Execute( query ) 
%>      
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
	      <meta http-equiv="REFRESH" content="5; url=Transferencias_Fondos_Proces.asp?count=<%= the_count-1%>&Systrace=<%=SYSTRACE%>&TDT=<%=TDT%>&Cta_debito=<%=Cta_Debito%>&Cta_credito=<%=Cta_Credito%>&importe=<%=Importe%>&describe=<%=describe%>">
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

   <%Else    'este es el else de     If rs2.EOF Then   %>
 
<% Cadena = UCASE(rs2("bit48resp")) 
 CADENA1 = rs2("bit48resp") 
 P=INSTR(CADENA,"REF_CORRIE:") 
 RC=mid(cadena, p+11,8) 
 P1=INSTR(CADENA,"FEC_CONTAB:") 
 FC=mid(cadena, p1+11,8) 

   query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rs2("respcode") & "')"
   set rs3 = conn.Execute( query3 ) 
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
      <td width="90%" align="center" colspan="2" height="46">
        <table border="0" cellpadding="5" cellspacing="0" width="100%">
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
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Fecha
    Contable:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<% = FC%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Cuenta
    Debitada:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<% =Cta_Debito %></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Referencia
    Corriente:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<% = RC%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Cuenta
    Acreditada:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<% =Cta_Credito %></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Importe
    Transferido:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>$ <% =FormatNumber(Importe,2) %></strong></font></td>
  </tr>
  <tr>
<%
   P = INSTR(CADENA,"PAGANDO") 
   fin = INSTR(CADENA,"COMPROBANTE")
   
   if (fin <> 0) and (p <> 0) then  
     describe = mid( CADENA, p, fin-p )
   end if   
   
   p = INSTR(CADENA,"REF_CORRIE:")
   
   if (fin <> 0) and (p <> 0) then
     ELECTR = mid( CADENA, fin+14, p-fin-14 )
   end if
%>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Observaciones:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<%=describe%></strong></font></td>
  </tr>
  <tr>
    <td width="40%" align="right" height="19"><font color="#000000" face="Verdana" size="2">Resultado
    de la Transferencia:</font></td>
    <td width="50%" align="left" height="19"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<% =rs3("spanish") %></strong></font>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right" height="18"><font color="#000000" face="Verdana" size="2">Comprobante
    de la Transferencia:</font></td>
    <td width="50%" align="left" height="18"><font color="#000000" face="Verdana" size="2"><strong>&nbsp;<%=ELECTR%></strong></font></td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="52">&nbsp;</td>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="18"><blockquote>
      <font color="#000000" face="Verdana" size="2"><strong><p align="left">Hecho:
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Autorizado:</p>
      </strong></font>
    </blockquote>
    </td>
  </tr>
  <tr>
<%If rs2("RESPCODE") <> "00" Then %>
<%If rs2("RESPCODE") = "01" Then %>
  </tr>
  <tr>
    <td width="90%" align="center" colspan="2" height="46"><strong><font face="Verdana" color="#FF0000">IMPORTANTE
          :</font><font color="#000000" face="Verdana"><small><br>
    </small>Esta transferencia no será tramitada hasta el siguiente día hábil, <small><br>
    </small>debido a que la sucursal de destino no está conectada a la Red Pública de
    Transmisión de Datos.</font></strong></td>
  </tr>
<%Else%>
  <tr>
    <td width="90%" align="center" colspan="2" height="46"><strong><font face="Verdana" color="#FF0000">IMPORTANTE
          :</font><font color="#000000" face="Verdana"><small><br>
    </small>Esta operación no se ejecutó correctamente.<small><br>
    </small>Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></strong></td>
  </tr> 
<%End If
 Else%>
  <tr>
    <td width="90%" align="center" colspan="2" height="18"><font color="#000000" face="Verdana" size="2"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
  </table>
  <Div id="B1" >
 	<p align="center">
      <a href="#" onClick="window.open 'Comprobante_Small.asp?FC=<%=FC%>&Cta_Debito=<%=Cta_Debito%>&RC=<%=RC%>&Cta_Credito=<%=Cta_Credito%>&Importe=<%=Importe%>&describe=<%=describe%>&Resultado=<%=rs3("spanish")%>&ELECTR=<%=ELECTR%>&Tipo=TRANSF','SubMenu','height=300,width=400,resizable,scrollbars,statusbar'"> Imprimir Comprobante</a></P>
      
      </div>
<%End If 
End If%>
</body>
</html>