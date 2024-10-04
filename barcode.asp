<%@ LANGUAGE = "JScript" %>
<%  Response.Expires=0; 
    Response.Buffer = true;
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
.style18 {
	font-family: Verdana;
	font-size: 14px;
	color: #FFFFFF;
}
.style1 {
	font-family: Verdana;
	font-weight: bold;
	color: #FFFFFF;
	font-size: 16px; 
	filter:glow(color=#000000,strength=2);
    width:100%;
}
-->
</style>

<script language="jscript">
function Validator()
{
  	
	var barcodeInfo = form1.barcode.value;
        form1.submit();
		
}

function detectEnter()
{
  if (event.keyCode == 13)
  {
	  Validator();
  }
}

function Autenticar()
  {     
	    form2.action = 'Autenticar.asp?User='+form2.user.value+'&Password='+form2.password.value;
        form2.submit();
  }
</script>


</head>

<body leftmargin="0" topmargin="0" style="overflow: hidden" onLoad="form1.barcode.focus()" onMouseDown="form1.barcode.focus()">

   <table width="100%" height="100%"  border="0" align="left" cellpadding="0" cellspacing="0" background="images/login_back.jpg">
     <tr align="center" valign="top" >
       <td width="40%" height="130" background="VSucursal/Images/banner.jpg" style="background-repeat:no-repeat"><img src="images/banner.jpg" width="1024" height="133"></td>
     </tr>
     <tr>
       <td align="center" valign="top">
	   
	   <table width="40%"  border="0" align="center" cellpadding="5" cellspacing="0" style="position:absolute; top:250px; left: 30%">
          <tr valign="top">
            <td colspan="3" height="47"><span class="style13"><img src="images/login.jpg" width="212" height="47" style="position:absolute; top:-5px; left: 5%"></span></td>
         </tr>
          <tr>
           <td colspan="3"><hr width="90%" size="1" color="#FFFFFF"></td>
         </tr>
          
		  <tr>
            <td colspan="3">
			<span class="style1">
			<%
			 Mensaje = Number(Request("Msg"));
			 Texto = "";
			 
			 switch(Mensaje)
			 {
			 case 0: Texto = "Tarjeta de Identificación incorrecta."; break;
			 case 1: Texto = "Ha superado el número máximo de intentos. <br> Favor de proporcionar el C&oacute;digo de Barras que aparece en su Tarjeta de Identificaci&oacute;n de Usuario utilizando el Lector de C&oacute;digos de Barras."; break;
			 case 2: Texto = "Su sesión ha expirado. Debe autenticarse nuevamente."; break;		
			 case 3: Texto = "Favor de proporcionar el C&oacute;digo de Barras que aparece en su Tarjeta de Identificaci&oacute;n de Usuario utilizando el Lector de C&oacute;digos de Barras."; Session.Abandon(); break;
			 default: Texto = "Favor de proporcionar el C&oacute;digo de Barras que aparece en su Tarjeta de Identificaci&oacute;n de Usuario utilizando el Lector de C&oacute;digos de Barras."; break;
			 }	
			 Response.Write(Texto);		
			 %>
			 </span>
			 </td>
          </tr>
          <tr align="center">
            <td colspan="3"><img src="images/lector.jpg" width="210" height="162" style="position:absolute; top:193px; left: -39">
			<img src="images/id.jpg" width="116" height="198" style="position:absolute; top:131px; left: 317">
			<img src="images/rayo.gif" width="337" height="102" style="position:absolute; top:196px; left: 86">
			</td>
          </tr>
		  
          <form name="form1" id="form1" method="post" action="login.asp">
	  
		  <tr>
            <td colspan="3">
				<input type="password" name="barcode" id="barcode" onkeypress="detectEnter()" onBlur="form1.barcode.focus()" style="width: 0%; border: 2 px; background-color:#E0C6AB; border-style:dashed; height: 30px;margin-left:0px;text-align:center; ">
			</td>
          </tr>
		  </form>
		         
       </table>
	   
	   </td>
     </tr>
     <tr align="center" valign="top">
       <td>&nbsp;</td>
     </tr>
   </table>
</body>
</html>