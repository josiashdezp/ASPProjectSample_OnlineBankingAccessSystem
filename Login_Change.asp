<%@ LANGUAGE = "JavaScript" %>
<% Response.Expires = 0
   Response.Buffer = true 
   if (String(Session("UsrId")) == "undefined")
      {Response.Redirect("../Main.asp");}
   %>
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="GENERATOR" content="Microsoft FrontPage 4.0">
   <title>Cambio de Contraseña</title>
</head>
<body text="#002299" bgcolor="#ffffff" link="#000066" vlink="#ff6600" alink="#003366" onLoad="document.form1.usrID.focus();">
<p><img border="0" src="../Images/Cambiar_Contrasena.gif" WIDTH="560" HEIGHT="50"></p>
<p>&nbsp;</p>
<% if( Request.Form("hname") != "hvalue" ) 
    { //Permite la captación de datos en la forma HTML 
    Session("RndVal") = Math.random().toString(); %>

<SCRIPT LANGUAGE="JavaScript" SRC="../Lib/md5.js" type="text/javascript"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../Lib/sha1.js" type="text/javascript"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
 var sharedValue = "<%=Session("RndVal")%>"
 function handleLogin() 
 { sendMD5Value() }
 
 function calculateMD5Value(pw) 
 { pw = hex_sha1(pw)
   pw = pw.toUpperCase() + sharedValue
   return hex_md5(pw)  }
 
 function sendMD5Value() 
 { document.forms["form1"].elements["password"].value = calculateMD5Value(document.forms["form1"].elements["password"].value)
   var
   pw0 = document.forms["form1"].elements["newPassword"].value
   document.forms["form1"].elements["PwLength"].value = pw0.length
   
   pw0 = hex_sha1(pw0)
   document.forms["form1"].elements["newPassword"].value = pw0.toUpperCase()
   
   var
   pw1 = hex_sha1(document.forms["form1"].elements["confirm"].value)
   document.forms["form1"].elements["confirm"].value = pw1.toUpperCase()
   document.forms["form1"].submit() }
</SCRIPT>
<center>
<div align="center">  <center>
<form METHOD="post" ACTION="Login_Change.asp" id="form1" name="form1" >
<table border="0" cellPadding="5" cellSpacing="1" width="65%">
<tr>
  <input type="hidden" name="hname" value="hvalue">
  <input type="hidden" name="RndValue" value="<%=Session("RndVal")%>">
  <input type="hidden" name="PwLength" value="0">
        <td bgColor="#FFFFFF" height="30">
            <%Expire=Request("Expire")
			 if( Expire=="True" ){ %>
              <div align="center"><b>
              <span style="background-color: #0000FF"><font face="Arial" size="3" color="#FFFFFF">
              <marquee>Su Contraseña ha expirado por lo que debe ser Cambiada.</marquee>
              </font></span>
              </b></div>
             <%}%>
</td>
    <tr>
        <td bgColor="#FFFFFF" align="left" vAlign="top" height="106" style="border-style: groove; border-color: #0000FF" colspan="2">
            <div align="right"><font face="Verdana" size="2" color="#000000"><b>Nombre de Usuario: </b></font><input name="usrID" size="14" maxlength="8" id="usrID" style="HEIGHT: 22px; WIDTH: 108px"> 
            </div>
            <div align="right">
            <div align="right"><font color="#000000" face="Verdana" size="2"><b>Contraseña
              Actual:</b></font> 
              
            <input id="password" name="password" style="HEIGHT: 22px; WIDTH: 108px" type="password"></font> 
            </div>
            <div align="right"><b><font face="Verdana" size="2" color="#000000">Nueva
              Contraseña:</font></b>
              <input id="newPassword" name="newPassword" style="WIDTH: 109px; EIGHT: 22px" type="password"></div>
              <font color="#000000" face="Verdana" size="2"><b>Confirme la nueva 
              Contraseña:</font> 
              <input id="confirm" name="confirm" style="HEIGHT: 22px; WIDTH: 109px" type="password"></div>
        <tr>
        <td bgColor="#FFFFFF" height="70" valign="middle" align="center" colspan="2">
<input type="submit" value="Cambiar" id="submit1" name="submit1" style="background-color: #000080; font-family: Verdana; color: #FFFFFF; font-weight: bold; border-style: outset; border-color: #000080" onClick="handleLogin()">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="reset" value="Deshacer" id="reset1" name="reset1" style="left: 163px; top: 1px; font-family: Verdana; background-color: #000080; color: #FFFFFF; font-weight: bold; border-style: outset; border-color: #000080">
    </tr></table>
	</form>
	</center>
</div>
</center>
<div align="center"><font face="Verdana" size="2"></font>&nbsp;</div>
<div align="center"><font face="Verdana" size="2" style="FLOAT: right; FONT-FAMILY: "><font color="black" size="4">
</font>
<p></p></div>

<%}
  else // si no es Request.Form("hname") != "hvalue"  entonces
  {
%>
<!--#include file="../Lib/md5.inc" -->
<% 
   function calculateMD5Value(pw) 
   { pw += Session("RndVal")
     return hex_md5(pw) }
   
   username 	  = Request.Form("usrID");
   clientPassword = Request.Form("password");
   
   var newPassword = new String( Request.Form("newPassword") );
   var confirm 	   = new String( Request.Form("confirm") );
   var PwLength    = Request.Form("PwLength")
      
   // Coger el password de la base de datos
   
   conn = Server.CreateObject( "ADODB.Connection" );
   conn.Open ("File Name="+Server.MapPath("../Connections/Conn_MSAcces.udl"));
   
   query = "Select * from Acceso Where (login='" + username + "')"
   rs = conn.Execute( query )   
   
   // Si el usuario existe en la base registro
   
   if( !rs.EOF )
   {  // es decir que existe ese usuario en la BD	  
     var 
	 bdpassword = new String(rs("Password"));
  
  	 serverPassword = calculateMD5Value(bdpassword);
     if ((username != Session("UsrId") && (String(Session("UsrId"))!="undefined")))
        {Response.Redirect("Login_Error.asp?Code=05");}
        else
  		  { if(clientPassword == serverPassword)
              { // El pasword anterior coincide
	            if( newPassword.valueOf() != confirm.valueOf() )
		          { // El nuevo password no coincide con su confirmacion
        	 		Response.Clear();
	     	 		Response.Redirect("Login_Error.asp?Code=01");
		   		  }
		        else
		          { if( PwLength < 8 )
			   		  {// El nuevo password es demasido corto
		        	   Response.Clear();
     				   Response.Redirect("Login_Error.asp?Code=02");
					  }
			   		else
				 	   { if( newPassword.valueOf() == bdpassword.valueOf() )
					 		{// El nuevo password es igual al password anterior
    	        	  		 Response.Clear();
	        	      		 Response.Redirect("Login_Error.asp?Code=03");
					  		}
					  	 else
					   	   {// Cambiar el password
		     			 	var hoy = new Date();
					     	var mes = hoy.getMonth()+1;
					     	query = "UPDATE Acceso SET Acceso.Password = '"+ newPassword +"', Acceso.Cambiar_password = False, Acceso.Fecha_Validacion = #"+ mes + "/" + hoy.getDate() + "/"+ hoy.getFullYear() +"# Where (Acceso.login='" + username + "');";
            			 	conn.Execute(query)   ;
		 			     	Response.Redirect("Login_Change_Success.htm");
		 				    }// cambiar password
				        } // else de Lenght<8
	  		       } // else de nuevos password no coinciden
              } // client password != serverpassword
	 		  else
	 		  {// El usuario existe pero el password anterior esta incorrecto
	    	   Response.Clear();
               Response.Redirect("Login_Error.asp?Code=04");
	           }
		 } // Session(UserId) = username
    } // RS.EOF, es decir el usuario noexiste en la BD.
 else
  {Response.Clear();
   Response.Redirect("Login_Error.asp?Code=05")
  }
}// fin de else (si no es Request.Form("hname") != "hvalue"  entonces)
%>
</body>
</html>