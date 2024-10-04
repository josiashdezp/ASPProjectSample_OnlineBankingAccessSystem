<%@LANGUAGE="JAVASCRIPT" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN"
"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<title>Sistema Automatizado Para el Control de las No conformidades DATADEF</title>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<meta name="description" content="Your description goes here." />
<meta name="keywords" content="your,keywords,goes,here" />
<meta name="author" content="Your Name / Original design by Andreas Viklund - http://andreasviklund.com" />
<link href="/Tesis/Styles/StyleDatadef.css" rel="stylesheet" type="text/css" />
</head>
<body>
     <h1>DATADEF</h1>
    <h2>Sistema Automatizado Para el Tratamiento de las No Conformidades.</h2>
    <div id="contentIndex" style="width:65%;margin-left:20%"> 
  <table width="95%" align="center" style="font:100% Verdana,Tahoma,Arial,sans-serif;text-align:left;">
    <tr> 
      <td colspan="2" style="border-bottom:1px solid;text-align:left;"><span class="darkNotice" style="font-size:2em;">Error  en tiempo de ejecuci&oacute;n: 
<%
   var ErrorType = "",
       ErrorMsg  = "",
	   ErrorNumber,
       Errores;
	   
		Errores = Server.GetLastError();
	    ErrorNumber = Errores.Number;
		ErrorType   = Errores.Category; 
		ErrorMsg    = Errores.Description;
%>
        </span></td>
    </tr>
    <tr> 
      <td style="text-align:left;"><span class="intro">Tipo de Error:</span> </td>
      <td style="text-align:left;"> <% Response.Write(ErrorType);%> </td>
    </tr>
    <tr style="text-align:left;"> 
      <td width="18%" style="vertical-align:top;"><span class="intro" >Descripci&oacute;n:</span></td>
      <td width="82%"> <%Response.Write(ErrorMsg);%> </td>
    </tr>
    <tr > 
      <td colspan="2" style="border-bottom:1px solid;text-align:center;"><%=Errores.Line%></td>
    </tr>
    <tr >
      <td colspan="2" style="border-bottom:1px solid;text-align:center;"><a href="javascript:window.history.back();" style="vertical-align:middle;">&lt;&lt; 
        Atrás</a></td>
    </tr>
  </table>
</div>
 </div>  
</body>
</html>