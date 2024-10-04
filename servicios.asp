<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<%  Response.Expires= 0 ;  
    Response.Buffer = true ;
	Response.CacheControl="no-cache" ;
	%>
<%
	
Destino = "";

	Sender = String(Request("Sender"));
	switch(Sender)
	{ case "Main": 
	  if (String(Request("Serv"))!="undefined")
	  {Destino = "Servicios/" + Request("Dest") + "?Serv=" + Request("Serv");}
	  else
	  {Destino = "Servicios/" + Request("Dest");}
	  break;
	  case "Out" : Response.Redirect("barcode.asp?Msg=3");	  				break;
	  default    : Destino  = "welcome.asp?Empresa=" + Request("Empresa");	break;
	}
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Sucursal Virtual</title>
<style type="text/css">
<!--
body {
	background-color: #9A2945;
}
.style13 {font-family: Verdana; font-size: 12px; font-weight: bold; color: #000000; }
-->
</style>
</head>
<body leftmargin="0" topmargin="0" style="overflow: hidden">
<%
  Whois = Session("UsrId");
  
  var Conn = Server.CreateObject("ADODB.Connection");
  Conn.ConnectionString = "File Name=" + Server.MapPath("./Connections/ConnString_BandecOnline.udl"); 

  Conn.Open; 
  query = "Select Servicios from tbl_Contratos Where (login='" + Whois + "')";
  rs1 = Conn.Execute( query );
  
  if (rs1.EOF) 
  		{
	     perm = 5695;  //Todos los servicios excepto List. oper., Lotes y O Cobro;
	    }	
  else
  {
    perm = String(rs1("Servicios")).toUpperCase();
	if (perm == 0)  perm=5695
  }
  
  
  /*
  
  'Estado de cuentas - 1
  'Disponibilidad    - 2 
  'Ultimos 10 Mov    - 4
  'Transferencias    - 8
  'Aporte			 - 16
  'Amortizacion		 - 32
  'Fincimex			 - 64
  'Listado Oper.	 - 128 
  'Lotes             - 256  
  'Comprobantes      - 512
  'Tipo de cambio    - 1024   
  'Envio de archivos - 2048
  'Sol. chequera     - 4096
  'Orden de Cobro    - 8192  
  'Transferencias de Fondos - Correo - 16384
  'Estado de cuentas - Correo - 32768
  '10 Mov. - Correo - 65536
  
  '16383 - Todos los servicios
  '7807  - Todos los servicios excepto List. oper., Lotes y O Cobro
  '8063  - Todos los servicios excepto List. oper. y O Cobro
  '7935  - Todos los servicios excepto Lotes y O Cobro
  '15999 - Todos los servicios excepto List. oper. y Lotes
  
  */
  %>	

<table width="100%" height="100%"  border="0" align="left" cellpadding="0" cellspacing="0" background="images/servicios_back.jpg">
  <tr align="center" valign="top" >
    <td width="700" height="130" colspan="2" background="VSucursal/Images/banner.jpg" style="background-repeat:no-repeat"><img src="images/banner.jpg" width="1024" height="133"></td>
  </tr>
  <tr>
    <td width="40%" align="center" valign="top"></td>
    <td width="60%" align="center"></td>
  </tr>
  <tr align="center" valign="top">
    <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td width="40%" align="center" valign="middle"><table width="80%"  border="0" align="center" cellpadding="2" cellspacing="0">
            <% if ((perm && 4) == 4)
			{%>
			<tr>
              <td>&nbsp;&nbsp;&nbsp;&nbsp;<a href="servicios.asp?Sender=Main&Dest=10_Mov.asp"><img  border="0" src="images/button_10mov.gif" onMouseOver="this.src='images/button_10mov_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_10mov.gif'" ></a></td>
            </tr>
            <%
			}
			
			if ((perm && 8) == 8) {
			%>
            <tr>
              <td><a href="servicios.asp?Sender=Main&Dest=Consultar_Sucursal.asp&Serv=TR"><img border="0" src="images/button_transferencias.gif" onMouseOver="this.src='images/button_transferencias_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_transferencias.gif'"></a></td>
            </tr>
            <%
			}
			if ((perm && 16) == 16) {
			%>
            <tr>
              <td><a href="servicios.asp?Sender=Main&Dest=Consultar_Sucursal.asp&Serv=AP"><img border="0" src="images/button_aportes.gif" onMouseOver="this.src='images/button_aportes_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_aportes.gif'"></a></td>
            </tr>
            <%
			}
			if ((perm && 32) == 32) {
			%>
            <tr>
              <td>&nbsp;&nbsp;&nbsp;&nbsp;<a href="servicios.asp?Sender=Main&Dest=Amortizacion.asp"><img border="0" src="images/button_amortizar.gif" onMouseOver="this.src='images/button_amortizar_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_amortizar.gif'"></a></td>
            </tr>
            <%
			}
			if ((perm && 1) == 1) {
			
			%>
            <tr>
              <td>&nbsp;&nbsp;&nbsp;&nbsp;<a href="servicios.asp?Sender=Main&Dest=Estado_Cuenta.asp"> <img  border="0" src="images/button_estadocuenta.gif" onMouseOver="this.src='images/button_estadocuenta_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_estadocuenta.gif'"></a></td>
            </tr>
			            <% }
			if ((perm && 4096) == 4096)
			{						%>
            <tr>
              <td>&nbsp;&nbsp;&nbsp;&nbsp;<a href="servicios.asp?Sender=Main&Dest=Consultar_Sucursal.asp&Serv=CH"> <img  border="0" src="images/button_chequeras.gif" onMouseOver="this.src='images/button_chequeras_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_chequeras.gif'"></a></td>
            </tr>
			 <%
			 }
			 if ((perm && 512) == 512)
			 {
			 
			 %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="servicios.asp?Sender=Main&Dest=Comprobante.asp"> <img  border="0" src="images/button_comprobante.gif" onMouseOver="this.src='images/button_comprobante_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_comprobante.gif'"></a></td>
            </tr>
            <%}%>
            <tr>
              <td align="center" valign="top"><a href="servicios.asp?Sender=Out"><img border="0" src="images/button_salir.gif" onMouseOver="this.src='images/button_salir_over.gif' ; this.style.cursor='hand'" onMouseOut="this.src='images/button_salir.gif'"></a></td>
            </tr>
          </table></td>
          <td width="60%"><iframe name="main" src="<% Response.Write(Destino);%>" width="95%" height="500" frameborder="0" allowtransparency="True" style="overflow:auto;"></iframe></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--

 
  
   
    
	 
	  
-->
</body>
</html


