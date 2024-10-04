 <html>
<head>
<title>Chequeo de Conexión</title>
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
			color : "#0000F0";
			font-size : 14px;		}

		A:hover		{
			text-decoration : underline;
			color : "#0000FF";		}
	</style>
</head>
<body style="background-color: transparent">
 <% the_count = CInt(Request("count")) 
   SYSTRACE = Request("Systrace")
   TDT = Request("TDT")
   Sucu = Request("Sucu")
   Serv = Request("Serv")
   set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")

   query2 = "SELECT bit48resp, respcode FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (respcode <> '') AND (Tdt = '" & TDT & "')"
   set rs2 = conn.Execute( query2 ) 
    If not rs2.EOF Then  'es decir si no hay respuesta
		If rs2("respcode")= 00 Then 
	    	Session("CheckConnectTime")=Time() 'Esta es la variable donde chequeo el tiempo para varificar conexion
			rs2.Close
			Select Case Serv
		      Case "TR"   Response.Redirect("Transferencias_Fondos.asp") '0 es Transferencias de Fondos
		      Case "AP"   Response.Redirect("Aporte.asp") '1 es Aporte al presupuesto
			  case "CH"   Response.Redirect("Chequera_Solicit.asp") '3 es solicitar chequera
   	       End Select
	 	Else
	   		query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rs2("respcode") & "')"
	    	set rs3 = conn.Execute( query3 ) 
	  		Response.Write(rs3("spanish"))
			rs3.Close 
	 	End If 
	Else
	  If the_count = 0 then  'y ya se acabaron los intentos se muestra el mensaje de error %>

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
	<%Else 'si no se han acabado los intentos seguimos tratando de conectarnos%>
	  <meta http-equiv="REFRESH" content="5; url=Consultar_Sucursal_Procesando.asp?count=<%= the_count-1%>&Systrace=<%=SYSTRACE%>&TDT=<%=TDT%>&Sucu=<%=Sucu%>&Serv=<%=Serv%>">

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
       <%End If	     'end de if count =0 %>

<% end if      'si el record set no está vacío
%>
 
</body>
</html>