<html>
<head>
<title>Listado de adeudos</title>
<style type="text/css">
<!-- 
.style1 {
	font-family: Verdana;
	font-size: 12px;
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
	font-size: 14px;
	
}
.style6 {
	font-family: Verdana;
	color: #000000;
	font-size: 16px;
	
}
.style11 {
	font-family: Verdana;
	font-weight: bold;
	color: #FFFFFF;
	font-size: 20px; 
	filter:glow(color=#000000,strength=2);
    width:100%;
}
-->
</style>
 <style>
 
 body {
	scrollbar-arrow-color: #DDCE67;
	scrollbar-base-color: #f0b6c4;
	scrollbar-face-color: #9f314e;
	scrollbar-highlight-color: #f0b6c4;
	scrollbar-shadow-color: #DDCE67;
}
 </style>

<style>
		A		{
			text-decoration : none;
			color : "#0000FF";
			font-size : 14px;		}

		A:hover		{
			text-decoration : underline;
			color : "#FF0000";		}
	.style12 {
	font-family: Verdana;
	font-size: 18px;
	font-weight: bold;
	color: #FF0000;
}
.style13 {
	font-family: Verdana;
	font-size: 16px;
	font-weight: bold;
	color: #FF0000;
}
.style14 {font-family: Verdana}
</style>
</head>
<body style="background-color: transparent" leftmargin="0" topmargin="0">
<% 
   the_count = CInt(Request("count")) 
   SYSTRACE = Request("Systrace")
   TDT = Request("TDT")

   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4)   

   On error resume next
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

          <meta http-equiv="REFRESH" content="5; url=Amort_Procesando_Adeudos.asp?count=<%= the_count-1%>&Systrace=<%=SYSTRACE%>&TDT=<%=TDT%>&Cta_credito=<%=Cta_Credito%>">
  <table width="70%" cellpadding="5" cellspacing="3" align="center">
    <tr bgcolor="#9C2E4B"> 
      <td colspan="3" align="center"><b><font face="Arial" color="white"> 
        <marquee scrollamount="5" width="100%" behavior="alternate">
        <font size="2" face="Verdana">...Procesando Adeudos...</font> 
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
		Else%>
<%If rs2("RESPCODE") = "00" Then 
     query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rs2("respcode") & "')"
     set rs3 = conn.Execute( query3 ) 
     
     CueMay = "CUE" & sucu
     query1 = "SELECT Nom_Client FROM " & Cuemay & " WHERE (Cod_Contra = '" & whois & "')"
     set rs1 = conn.Execute( query1 )
     empresa = rs1("Nom_client")
  %>
<TABLE width="100%" border=0 align=center cellPadding=5 cellSpacing=2>
  <TR bgcolor="#FFCCCC">
    <TD width="10%" align=right valign="top" style="WIDTH: 10%">
     <span class="style3">Sucursal:</span>
    </TD>
    <TD width="50%" align=left valign="top" style="WIDTH: 30%">
    <span class="style3"><%=sucu%></span>    </TD>
    <TD width="20%" align=right valign="top" style="WIDTH: 20%">
     <span class="style3">Fecha  Emisión:</span>
    </TD>
    <TD width="20%" height=20 align="left" valign="top" style="HEIGHT: 20px; WIDTH: 20%">
    <span class="style3"><%=date%></span>    </TD>
  </TR>
  <TR bgcolor="#FFCCCC">
    <TD colspan="2" align=left style="WIDTH: 10%">
     <span class="style3"><%=empresa%></span>    </TD>
    <TD align=right valign="top" style="WIDTH: 20%">
   <span class="style3">Hora  Emisión:</span>
    </TD>
    <TD align="left" valign="top" style="WIDTH: 20%">
    <span class="style3"><%=time%></span></TD>
  </TR>
</TABLE>
<BR>

<TABLE align=center border=0 cellPadding=5 cellSpacing=2 width="100%" style="LEFT: 10px" borderColor=#000000 borderColorDark=#000000>
    <TR>
      <TD colspan="3" align=middle vAlign=center bgColor="#FFCCCC" style="WIDTH: 20%">&nbsp;</TD>
      <TD colspan="4" align=middle vAlign=center bgColor=#FCE8AB style="WIDTH: 15%"><font color="#000000" face="Verdana" size="2"><strong>Próxima Amortización</strong></font> </TD>
    </TR>
    <TR>
    <TD align=middle style="WIDTH: 20%" vAlign=center width="20%" bgColor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><strong>Cuenta</strong></font>
    </TD>
    <TD align=middle style="WIDTH: 10%" vAlign=bottom width="10%" bgColor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><strong>Cod. Obj.</strong></font>
    </TD>
    <TD align=middle style="WIDTH: 10%" vAlign=bottom width="10%" bgColor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><strong>Cod. Inv.</strong></font>
    </TD>
    <TD align=middle vAlign=center style="WIDTH: 15%" width="15%" bgColor=#FCE8AB>
      <font color="#000000" face="Verdana" size="2"><strong>Saldo</strong></font>
    </TD>
    <TD align=middle vAlign=center style="WIDTH: 15%" width="15%" bgColor=#FCE8AB>
      <font color="#000000" face="Verdana" size="2"><strong>Fecha</strong></font>
    </TD>
    <TD align=middle vAlign=center style="WIDTH: 15%" width="15%" bgColor=#FCE8AB>
      <font color="#000000" face="Verdana" size="2"><strong>Principal</strong></font>
    </TD>
    <TD align=middle vAlign=center style="WIDTH: 15%" width="15%" bgColor=#FCE8AB>
      <font color="#000000" face="Verdana" size="2"><strong>Intereses</strong></font>
    </TD>
  </TR>
 
  <% Cadena = UCASE(rs2("bit48resp")) 
   CADENA1 = rs2("bit48resp")   
  

     TotalSaldo = 0
     TotalProx = 0
     TotalInt = 0
     p = instr(CADENA,"CUENTA:") 
   do while p <> 0 
  
 Cuenta = mid(cadena, p+7, 14) 
  
  p = instr(p,CADENA,"OBJCRE:") 
  if p <> 0 then 
       p1 = instr(p,CADENA,"INVCRE:")    
       ObjCre = mid(cadena, p+7, p1-p-7)      
     end if
  
  p = instr(p,CADENA,"INVCRE:") 
 if p <> 0 then 
       p1 = instr(p,CADENA,"SALDO:")    
       InvCre = mid(cadena, p+7, p1-p-7)      
     end if

  p = instr(p,CADENA,"SALDO:") 
 if p <> 0 then 
       Saldo = mid(cadena, p+6, 12)/100 
       TotalSaldo = TotalSaldo + Saldo
     end if

   p = INSTR(p,CADENA,"VENCTO:")
   if p <> 0 then 
       Vencto = mid(cadena, p+7, 8)
       dia = mid(Vencto, 7, 2)
       mes = mid(Vencto, 5, 2)
       ano = mid(Vencto, 1, 4)
  
       Vencto = dia & "/" & mes & "/" & ano
     end if

 p = INSTR(p,CADENA,"PROXIMA:") 
  if p <> 0 then 
       Proxima = mid(cadena, p+8, 12)/100 
       TotalProx = TotalProx + Proxima
     end if

   p = INSTR(p,CADENA,"INTERES:") 
  if p <> 0 then
       Interes = mid(cadena, p+8, 12)/100 
       TotalInt = TotalInt + Interes
     end if
  %>  
 
  <TR>
    <TD align=middle width="20%" bgcolor="#FFCCCC">
    <% 'Solo se pueden amortizar las 141x y 151x, no se permiten 1514 ni objeto de credito 1800 
	 if (ObjCre <> "1800") and ((mid(Cuenta,4,3)="141") or (mid(Cuenta,4,3)="151")) and (mid(Cuenta,4,4)<>"1514") then %>
      <font color="#000000" face="Verdana" size="2"><a target="main" href="Amortizar.asp?Cta_Credito=<%=Cuenta%>&Importe=<%=Proxima%>&Interes=<%=Interes%>&Tope=<%=Saldo+Interes%>"><b><%=Cuenta%></b></a></font>
    <% else %>
      <font color="#000000" face="Verdana" size="2"><b><%=Cuenta%></b></font>      
    <% end if %>
    </TD>
    <TD align=middle style="WIDTH: 10%" width="10%" bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=ObjCre%></small></font>
    </TD>
    <TD align=middle style="WIDTH: 10%" width="10%" bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=InvCre%></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=FormatNumber(Saldo,2)%></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=Vencto%></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=FormatNumber(Proxima,2)%></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=FormatNumber(Interes,2)%></small></font>
    </TD>
  </TR>
  <% p = instr(p,CADENA,"CUENTA:")
    loop %>  
   <TR>
    <TD align=middle width="20%" bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"></font>
    </TD>
    <TD align=middle style="WIDTH: 10%" width="10%" bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><strong>Total:</strong></font>
    </TD>
    <TD align=middle style="WIDTH: 10%" width="10%" bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=FormatNumber(TotalSaldo,2)%></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=FormatNumber(TotalProx,2)%></small></font>
    </TD>
    <TD style="WIDTH: 15%" width="15%" align=middle bgcolor="#FFCCCC">
      <font color="#000000" face="Verdana" size="2"><small><%=FormatNumber(TotalInt,2)%></small></font>
    </TD>
  </TR>
</TABLE>
<% Else
If rs2("RESPCODE") = "79" Then %>
 <table width="100%">
  <tr>
    <td width="90%" height="46" colspan="2" align="center" bgcolor="#FCE8AB">
     <strong>
       <font face="Verdana" color="#FF0000">IMPORTANTE:</font>
       <font color="#000000" face="Verdana">
       <small><br>
      Usted no tiene cuentas de préstamo. Por favor consulte con su sucursal antes de repetir esta operación.</small></font>     </strong>    </td>
  </tr>
</table>
  <%Else%> 
  <table width="100%">
  <tr>
    <td width="90%" height="46" colspan="2" align="center" bgcolor="#FCE8AB"><strong><font face="Verdana" color="#FF0000">IMPORTANTE:</font><font color="#000000" face="Verdana"><small><br>
    </small>Esta operación no se ejecutó correctamente.<small><br>
    </small>Es altamente recomendable que se verifique en la sucursal antes de repetirla.</font></strong></td>
  </tr> 
</table>
  <%End If
  End If
  End If%>
</body>
</html>