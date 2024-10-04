<%@ LANGUAGE = "JScript" %>
<%  Response.Expires=0; 
    Response.Buffer = true;

    if ((String(Request("barcode")) != "undefined")) // es decir si el campo barcode existe. eso es porque viene de la pagina barcode.
		{
		    codigo = String(Request("barcode"));
			
			conn = Server.CreateObject("ADODB.Connection");
		    conn.ConnectionString = "File Name="+Server.MapPath("Connections/ConnString_BandecOnline.udl");
		    conn.Open;
		    query = "Select Emp_Nombre from tbl_contratos Where (CodigoAcceso='" + codigo + "');";
		    rsBarCode = conn.Execute( query )   ;			
			
			if (rsBarCode.EOF)
			{Response.Redirect("barcode.asp?Msg=0");}
			Session("LogIn") = 0;
		}	

if( Request.Form("hname") == "hvalue" )
{  
		codigo  = String(Request("FormBarCode"));
		 
%>
<!--#include file ="Lib/md5.inc"-->
<%  function calculateMD5Value(pw) 
   {
     pw += Session("RndVal");
     return hex_md5(pw);
   }
   
   barras = codigo;
   username = Request.Form("login");
   clientPassword = Request.Form("pass");
     
   // Coger el password de la base de datos
   
   conn = Server.CreateObject("ADODB.Connection");
   conn.ConnectionString = "File Name="+Server.MapPath("Connections/ConnString_BandecOnline.udl");
   conn.Open;
   query = "Select * from tbl_contratos Where (login='" + username + "') and (CodigoAcceso = '"+ barras +"');";
   rs = conn.Execute( query )   
   
   // Si el usuario existe en la base registro
   
   if(!rs.EOF)
   {   
     var bdpassword = rs("Password");
	 serverPassword = calculateMD5Value(bdpassword);
	 	 
	if(clientPassword == serverPassword)
     {  
	 	Mi_Fecha = new Date( rs("Fecha_Validacion") )
        hoy = new Date()
		 	 
      if( (((hoy - Mi_Fecha)/86400000) < 90) && !(rs("Cambiar_Password")==true)) 
       {
     	 Response.Clear
	     Response.Redirect("Login_Redir.asp?login="+username)
       }
       else // La contrasena expiro o debe ser cambiada
       {
         Response.Clear
         Response.Redirect("Login_Change.asp?Expire=True")
       }
     }
     else  // La contrasena no coincide
     {
	    //Verificar si el passwd esta sin encriptar en la base registro
		AspObj = Server.CreateObject("Comp.CompCrypt") 
        cpassword = AspObj.Firmar( bdpassword )
		 
		serverPassword = calculateMD5Value(cpassword)

		// Si coincide el cliente encript con el Servidor encript guardo la del servidor encriptado.
		if(clientPassword == serverPassword) 
		{
		   hoy = new Date()
		   
		   dia = hoy.getDate();
		   mes = hoy.getMonth()+1;
		   anio = hoy.getFullYear();
		   
		   fecha = mes + "/" + dia+ "/" + anio;
		
		   
		   query = "UPDATE tbl_Contratos SET tbl_Contratos.Password = '"+ cpassword +"', tbl_Contratos.Fecha_Validacion = Convert(smalldatetime,'"+ fecha +"',101) Where (tbl_contratos.login='" + username + "')";
		   rs = conn.Execute(query)   
		   
           Response.Clear
	       Response.Redirect("Login_Redir.asp?login="+username)
		}
		else
		{ 
		  //  Se enciprtó el de la BD pero no coincide con el del cliente 
		  Session("LogIn")++;
		}
     }
   }
   else
   {
	 // El usuario Login con las barras esas no existen 
	Session("LogIn")++; 
   }  
 if (Session("LogIn") == 3) Response.Redirect("barcode.asp?Msg=1"); 
  } 
else
{
   Session("RndVal") = Math.random().toString();     
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
	body 	{background-color: #9A2945;		 }
	.style13 {font-family: Verdana; font-size: 12px; font-weight: bold; color: #000000; }
	.style18 {font-family: Verdana;	font-size: 14px;	color: #FFFFFF;}
	-->
</style>
<script language="jscript">

function detectEnter()
{
  if (event.keyCode == 13)
  {   Validator(); }
}

</script>
<SCRIPT LANGUAGE="JavaScript" SRC="Lib/md5.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="Lib/sha1.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">

 var sharedValue = "<% =Session("RndVal") %>"
 
 function handleLogin() 
 {
    sendMD5Value(calculateMD5Value());
 } 
 
 function calculateMD5Value() 
 {
   var pw = document.forms["form1"].elements["pass"].value
   pw = hex_sha1(pw)
   pw = pw.toUpperCase() + sharedValue
   return hex_md5(pw)
 }
 
 function sendMD5Value(hash) 
 {
   document.forms["form1"].elements["pass"].value = hash
 }
 
 function GoToPswField()
 {
	 Login 		= new String(document.forms["form1"].elements["login"].value);
	 Largo 	= Number(Login.length);
	 
	// if (Largo == 8) document.forms["form1"].elements["pass"].focus();
 }
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" style="overflow: hidden" onLoad="form1.login.focus()">
<form METHOD="post" ACTION="Login.asp" id="form1" name="form1" autocomplete="off">
  <input type="hidden" name="FormBarCode" value="<%=codigo%>">
  <table width="100%" height="100%"  border="0" align="left" cellpadding="0" cellspacing="0" background="images/login_back.jpg">
    <tr align="center" valign="top" >
      <td width="40%" height="130" background="VSucursal/Images/banner.jpg" style="background-repeat:no-repeat"><img src="images/banner.jpg" width="1024" height="133"></td>
    </tr>
    <tr>
      <td align="center" valign="top"><table width="40%"  border="0" align="center" cellpadding="5" cellspacing="0" style="position:absolute; top:250px; left: 30%">
          <tr valign="top">
            <td colspan="3" height="47"><span class="style13"><img src="images/login.jpg" width="212" height="47" style="position:absolute; top:-5px; left: 5%"></span></td>
          </tr>
          <tr>
            <td colspan="3"><hr width="90%" size="1" color="#FFFFFF"></td>
          </tr>
          <tr>
            <td colspan="3" align="center"><span class="style18"><b>Introduzca su Código de Usuario y Contraseña. </b></span></td>
          </tr>
          <tr>
            <td colspan="3"><hr width="90%" size="1" color="#FFFFFF"></td>
          </tr>
          <tr>
            <td width="20%" rowspan="3" align="center"><img src="images/login.gif" width="90" height="90" border="0"></td>
            <td width="35%" height="30" align="left" valign="middle" background="images/fondo_boton.gif" style="background-repeat:no-repeat; background-position: center left"><span class="style13">&nbsp;Usuario</span></td>
            <td width="35%" background="images/fondo_boton.gif" style="background-repeat:no-repeat; background-position: center left"><span class="style13">&nbsp;Contrase&ntilde;a</span></td>
          </tr>
          <input type="hidden" name="hname" value="hvalue">
          <input type="hidden" name="RndValue" value="<%=Session("RndVal")%>">
          <tr>
            <td height="30" background="images/fondo_boton1.gif" style="background-repeat:no-repeat; background-position: center left"><input type="text" name="login" id="login" style="width:75%;border:0px;background-color:#E0C6AB;" onKeyDown="GoToPswField()" >
            </td>
            <td height="30" background="images/fondo_boton1.gif" style="background-repeat:no-repeat; background-position: center left;vertical-align:middle;"><input type="password" name="pass" id="pass" style="width: 75%; border: 0 px; background-color:#E0C6AB">
            </td>
          <tr>
            <td height="30" colspan="3" align="center"><input type="submit" name="button" value="Entrar" onClick="handleLogin()"></td>
          </tr>
        </table></td>
    </tr>
    <tr align="center" valign="top">
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
