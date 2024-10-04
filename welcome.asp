<%  Response.Expires=0 
    Response.Buffer = true
    login = Request("Empresa")

  
  'El código que sigue es el mismo de INFORMA.ASP pero en este caso en vez de hacer un include
  'lo ponemos directamente.
  ' El código que está en informa se ejecuta cada vez que se llame cada uno de los servicios y está configurado para esto.
  
  
  whois = Session("UsrId")
  
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open "File Name=" & Server.MapPath("Connections/ConnString_BandecOnline.udl")
  query1 = "Select Cambiar_Password, Fecha_Validacion from tbl_Contratos Where (login='" & Whois & "')"
  

    set rs1 = Conn.Execute( query1 )
	
	if not rs1.eof then
	
    Mi_Fecha = rs1("Fecha_Validacion")
    If not rs1("Cambiar_Password")=True AND DateDiff("d", Mi_Fecha, Date()) < 90 then
      date_conn = Date()
	  time_conn = Time()
	  
      position = inStr( Request.ServerVariables("URL"),"/")
	  service=mid(Request.ServerVariables("URL"),position+11)
  
	  query = "INSERT INTO Informa (Tipo_Autenticacion, Usuario, Dir_Remota, Fecha_Conexion, Hora_Conexion, Servicio) VALUES ('" & Request.ServerVariables("AUTH_TYPE") & "',  '" & Session("UsrId") & "', '" & Request.ServerVariables("Remote_Addr") & "', " & "convert(smalldatetime,'" & Date() & "',101)" & ", '" & Time() & "', '" & service & "')"
	  
	  Conn.Execute( query )
	 Else
       Response.Clear
       Response.Redirect("change.asp?Expire=True")
     End IF  
    
	 else
	 end if

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Sucursal Virtual</title>
<style>
		A
		{
			text-decoration : underline;
			color : "#FFFFFF";
			
        }

		A:hover
		{
			text-decoration : underline;
			color : "#DDCE67";

		}
	</style>
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
}
.style3 {color: #FFFFFF;
font-size: 20px;
}
.style4 {
	font-family: Verdana;
	color: #FFFFFF;
	font-size: 18px;
	filter:glow(color=#000000,strength=2);
width:100%;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" style="overflow: hidden" style="background-color: transparent;">
<table width="100%%"  border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td><span class="style1">Bienvenido al Sistema.</span></td>
  </tr>
  <tr>
    <td><hr size="1" color="#FFFFCC"></td>
  </tr>
  <tr>
    <td><span class="style2">Usted esta identificado como: &nbsp;
     <b> <% Response.Write Request("Empresa") %> </b>
    </span></td>
  </tr>
  <tr>
    <td><hr size="1" color="#FFFFCC"></td>
  </tr>
  <tr>
    <td><span class="style4"><span class="style3">&#8226;</span> Haga Click sobre el servicio deseado en el Panel Izquierdo. </span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><span class="style2"><span class="style3">&#8226;</span> Para cualquier duda, queja o sugerencia contacte con el Personal Encargado en la sucursal. </span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><span class="style4"><span class="style3">&#8226;</span> &iquest;Alg&uacute;n Problema al utilizar los Servicios Online? Consulte nuestra <a href="Ayuda.asp">P&aacute;gina de Ayuda </a> donde encontrar&aacute; respuesta a los problemas m&aacute;s frecuentes... </span></td>
  </tr>
    <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><span class="style2"><span class="style3">&#8226;</span> No olvide Cerrar Sesi&oacute;n al terminar su trabajo. Para ello haga Click sobre el Bot&oacute;n Salir. </span></td>
  </tr>
  <tr>
    <td><hr size="1" color="#FFFFCC"></td>
  </tr>
</table>

</body>
</html>
