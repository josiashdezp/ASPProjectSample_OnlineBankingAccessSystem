<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Entrada al Sistema</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body BGCOLOR="#ffffff" topmargin="5">
<table width="100%" border="0" cellpadding="5" cellspacing="5">
  <tr>
      <td bgcolor="#3366CC"><font color="#FFFFFF" size="2" face="Arial"><b>Acceso al Sistema. </b></font></td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#C6E2FF">
            <font face="Arial" size="2">Para utilizar los servicios en línea que ofrecemos primero debe 
            identificarse como usuario del sistema. Para ello  
            teclee su <b>C&oacute;digo de Usuario</b> y <b>Contraseña.</b></font></td>
  </tr>
</table>
<table border="0" cellPadding="1" cellSpacing="1" width="100%" align="center" height="90">
    
    <tr>
        <td height="58" align="center"><img alt src="../Images/AccesoDenegado.gif" WIDTH="45" HEIGHT="44"></td>
        <td height="58" align="left"><font face="Verdana" color="#000000"><b>
		<% var Code = Request("Code");
		   Response.Write("Error número: <span style='color:red;font-weight:bold;'>"+Code+"</span>.<br>"); 
		   switch(Number(Code))
		   {case 01: Response.Write("La nueva contrase&ntilde;a y su confirmación no coinciden. Por favor, intentente de nuevo."); break;
			case 02: Response.Write("La nueva contraseña es demasiado corta, debe contener al menos 8 caracteres."); break;
			case 03: Response.Write("La nueva contrase&ntilde;a coincide con la contrase&ntilde;a que esta utilizando actualmente. Por favor, escoja otra."); break;
		   	case 04: Response.Write("Error de nombre o contraseña. Intente otra vez.");       break;// La contrase&ntilde;a actual esta incorrecta. Por favor, intente de nuevo.
			case 05: Response.Write("Error de nombre o contraseña. Intente otra vez.");       break;// Ud. no es Usuario del Sistema o el nombre de usuario no coincide con el de sus sesión. Contáctenos.
			case 06: Response.Write("Error de nombre o contraseña. intente otra vez.");       break; // acceso denegado.usted no es usuario del sistema
          }%>
		</b></font></td>
  </tr></table>
<div align="center"><strong><font color="#0000ff" face="Verdana" size="2"> 
   <font color="#0000ff" face="Verdana" size="2">
   <%if (Code==7) 
      {Response.Redirect("../Main.asp");}
    else 
   {%>
   [</font></font></font>
</font><font color="#0000ff"> </font><font color="black" face>
<a href="Login_Change.asp"><font color="blue" face="Verdana" size="2">Cambiar Contrase&ntilde;a</font></a></font><font color="blue"> </font> 
   <font color="#0000ff" face="Verdana" size="2">]
   <%}%>
   </font></strong> 
</div>
<table width="100%">
	<tr align="center">
	  <td colspan="2">
      <img src="../Images/line_ondulante_small.gif" width="350" height="38"> </td>
	<tr>
</TABLE>
</body>
</html>