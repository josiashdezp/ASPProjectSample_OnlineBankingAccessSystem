<HTML>
<HEAD>
<title>Aporte al presupuesto</title>
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
</HEAD>
<BODY style="background-color: transparent">
<% 
   the_count = CInt(Request("count")) 
   SYSTRACE = Request("Systrace")
   TDT = Request("TDT")
   
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4)         
   Cuemay = "CUE" & sucu
   
   Cta_Debito = Request("Cta_Debito")
   Importe = Request("Importe")
   set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 

   query2 = "SELECT bit48resp, respcode FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (respcode <> ' ') AND (Tdt = '" & TDT & "')"
   set rs2 = conn.Execute( query2 ) 
   If rs2.EOF Then
      If the_count = 0 then %>
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
        <meta http-equiv="REFRESH" content="5; url=Aporte_Procesando.asp?count=<%= the_count-1%>&Systrace=<%=SYSTRACE%>&TDT=<%=TDT%>&Cta_debito=<%=Cta_Debito%>&Importe=<%=Importe%>">             
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

     <%End If       	     	     	     	     	     	     	     	     
   Else
     set conn = Server.CreateObject( "ADODB.Connection" )
     conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
	 query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rs2("respcode") & "')"
     set rs3 = conn.Execute( query3 ) 
     
     query2 = "SELECT bit48resp, respcode FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (respcode <> ' ') AND (Tdt = '" & TDT & "')"
     set rs2 = conn.Execute( query2 ) 
  
     CADENA = rs2("bit48resp") 
     
	 p = instr( CADENA, "NIT:" )
     p1 = instr( p, CADENA, ";" )
     NIT = mid( CADENA, p+4, p1-p-4 )
     
	 p = instr( CADENA, "SUC:" )
     p1 = instr( p, CADENA, ";" )
     SUC = mid( CADENA, p+4, p1-p-4 )
     
	 p = instr( CADENA, "PF:" )
     p1 = instr( p, CADENA, ";" )
     Parrafo = mid( CADENA, p+3, p1-p-3 )
     
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
    
     p = instr( CADENA, "Comprobante:" )
     l = mid( Cadena, p+12, 2)
     Firma= mid( CADENA, p+14, l )
    
     p = instr( CADENA, "REF_CORRIE:" )
     RC = mid( CADENA, p+11, 8 )
    
	 p = instr( CADENA, "FEC_CONTAB:" )
     FC = mid( CADENA, p+11, 8 ) %>

<table border="1" width="100%" cellpadding="0" bordercolor="#000000" bordercolordark="#000000" cellspacing="0" align="center">
  <tr>
    <td width="90%" align="center" colspan="9" height="39">
      <table border="0" cellpadding="5" cellspacing="2" width="100%">
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
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Pago&nbsp; </strong></font><input type="checkbox" name="C1" value="ON" checked>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;<font face="Verdana"><%=NIT%></font></td>
  </tr>
  <tr><% Cuemay = "CUE" & sucu %>
      <td width="58%" align="right" height="31" colspan="2"><p align="left"><font color="#000000" face="Verdana" size="1">Debítese a:<br>
<%query = "SELECT Nom_Client FROM "& Cuemay &" WHERE (Cod_Contra = '" & whois & "')"
set rs = conn.Execute( query ) %>    </font><font color="#000000" face="Verdana" size="2"><%Response.Write(rs("Nom_Client"))
 Nom_Client = rs("Nom_Client")%></font><font color="#000000" face="Verdana" size="1"> </font></td>
    <td width="32%" align="center" height="31" colspan="7"><font color="#000000" face="Verdana"><font size="1">Código de la Cuenta:<br>
    </font><% =Cta_Debito%></font></td>
  </tr>
  <tr>
    <td width="58%" align="right" height="31" colspan="2"><p align="left"><font color="#000000" face="Verdana" size="1">Tributo:<br>
    </font><font color="#000000" face="Verdana" size="2"><%=mid(Parrafo,8)%></font><font color="#000000" face="Verdana" size="1"> </font></td>
    <td width="32%" align="center" height="31" colspan="7"><font color="#000000" face="Verdana" size="1">Código:<br>
    </font><font face="Verdana"><% =mid(PARRAFO,1,7) %></font></td>
  </tr>
  <tr>
    <td width="58%" align="right" height="38" colspan="2"><p align="left"><font color="#000000" face="Verdana" size="1">Importe:<br>
    </font></td>
    <td width="32%" align="center" height="38" colspan="7">
      <p align="center"><b>$</b>&nbsp; <font color="#000000" face="Verdana"><%=FormatNumber(Importe,2)%></font></p>
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
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Resultado:&nbsp;</small><font size="2"><b><% = rs3("spanish") %></b></font></font></td>
  </tr>
  <%If sucu <> SUC Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Tramitado por correo electrónico en la sucursal:&nbsp;</small><font size="2"><b><%=SUC%></b></font></font></td>
  </tr>
  <%End if
  If (rs2("RESPCODE") = "00") or (rs2("RESPCODE") = "01") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Referencia:&nbsp;</small><font size="2"><b><% = RC %></b></font></font></td>
  </tr>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><font face="Verdana"><small>Fecha contable:&nbsp;</small><font size="2"><b><% =FC %></b></font></font></td>
  </tr>
  <%END IF%>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="24" valign="top"><p align="center"><font face="Verdana"><small>Comprobante:
      </small></font><font color="#000000" face="Verdana"><b><% =Firma %></b></font></td>
  </tr>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="37"><blockquote>
      <font color="#000000" face="Verdana"><strong><p align="left">Hecho:
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Autorizado:</p>
      </strong></font>
    </blockquote>
    </td>
  </tr>
  <%If (rs2("RESPCODE") = "00") or (rs2("RESPCODE") = "01") Then %>
  <tr align="center">
    <td width="90%" align="center" colspan="9" height="18"><font color="#000000" face="Verdana"><strong>Para cualquier reclamación presente este comprobante.</strong></font></td>
  </tr>
  <%End If
  If rs2("RESPCODE") <> "00" Then 
     If rs2("RESPCODE") = "01" Then %>
       <tr align="center">
         <td width="90%" align="center" colspan="9" height="24" valign="top">
         <strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br>
         </small>Esta transferencia no será tramitada hasta el siguiente día hábil, <small><br>
         </small>debido a que la sucursal de destino no está conectada a la Red Pública de
         Transmisión de Datos.</font></strong>
         </td>
       </tr>
     <%Else%>
       <tr align="center">
       <td width="90%" align="center" colspan="9" height="24" valign="top">
       <strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br>
        </small>Esta operación no se ejecutó correctamente.<small><br>
        </small>Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></strong>
       </td>
       </tr> 
     <%End If
	 End If%>
</table>
 <Div id="B1" >
 	<p align="center">
    	<A href="#" onClick="window.open 'Comprobante_Small.asp?Tipo=APORT&FC=<%=FC%>&Cta_Debito=<%=Cta_Debito%>&RC=<%=RC%>&Importe=<%=Importe%>&Resultado=<%=rs3("spanish")%>&Parrafo=<%=Parrafo%>&NIT=<%=NIT%>&Nom_Client=<%=Nom_Client%>&Pcpal=<%=Pcpal%>&Recargo=<%=Recargo%>&Multa=<%=Multa%>&PD=<%=PD%>&PH=<%=PH%>&HD=<%=HD%>&TP=<%=TP%>&RF=<%=RF%>&II=<%=II%>&TI=<%=TI%>&IO=<%=IO%>&SUC=<%=SUC%>&Firma=<%=Firma%>','SubMenu','height=300,width=400,resizable,scrollbars,statusbar'">
        	<< IMPRIMIR COMPROBANTE >>
         </a>
    </P>
 </div>
 <%End If%>
</body>
</html>