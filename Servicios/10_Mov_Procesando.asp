<html>
<head>
<title>Ultimos 10 Movimientos</title>
 </head>
<body style="background-color: transparent ">
<% the_count = CInt(Request("count")) 
   SYSTRACE = Request("Systrace")
   TDT = Request("TDT")
   Cuenta = Request("Cuenta")
   On Error Resume Next
 
  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")

   query2 = "SELECT bit48resp, respcode FROM M_MENSWI WHERE (Systrace = '" & SYSTRACE & "') AND (respcode <> ' ') AND (Tdt = '" & TDT & "')"
   set rs2 = conn.Execute( query2 ) 
    If rs2.EOF Then
      If the_count = 0 then %>

<table border="0" cellSpacing="1" width="90%">
  <tr>
    <td align="middle" vAlign="center" width="50%"><font color="#dee30f"><IMG border=0 height=37 id=IMG1 src="../Images/Warning.gif" width=39></font></td>
    <td align="left" rowSpan="2" vAlign="center" width="50%"><FONT size=2><b><font color="#000000" face="Arial" >El tiempo requerido para 
      esta transacción ha sido excedido y no se ha podido establecer conexión con la 
      sucursal.</font></b> </FONT> 
      <p><b><font color="#000000" face="Arial" size="2">Por favor, intente de nuevo o 
      consulte con el Banco.</font></b></p></td></tr>
  <tr>
    <td align="middle" vAlign="center" width="50%"><IMG alt ="" border=0 height=20 hspace=0 src="../Images/Tiempo%20Excedido.gif" style    ="HEIGHT: 20px; WIDTH: 182px" useMap="" width=182 ></td></tr></table>
    
       <%Else%>          
	  <meta http-equiv="REFRESH" content="5; url=10_Mov_Procesando.asp?count=<%= the_count-1%>&Systrace=<%=SYSTRACE%>&TDT=<%=TDT%>&Cuenta=<%=Cuenta%>">
	
       <table width="60%" align="center" cellpadding="5" cellspacing="5">
    <tr bgcolor="#990033"> 
      <td colspan="3" align="center"><b><font face="Arial" color="white"> 
        <marquee scrollamount="5" width="100%">
        <font size="2" face="Verdana">COMUNICANDO CON LA SUCURSAL 
        <% = mid(Session("UsrId"), 1 , 4) %>
        ...</font> 
        </marquee>
      </font></b></td>
    </tr>
</table>
	        <%End If%>
<%Else
 If rs2("respcode")= 00 Then %> 

<% Cadena = UCASE(rs2("bit48resp"))%>

<table width="70%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr bgcolor="#FFD6D6">
    <td width="67%" align="middle" colspan="2">
	<strong><font face="Verdana"><font color="#000000">Cuenta:</font> <font color="#000000">
    <% = Cuenta %></font></font></strong></td>
  </tr>
  <tr>
    <td width="33%" align="middle" bgcolor="#FFD6D6"><font face="Verdana" color="#000000"><small><strong>Fecha
    Contable</strong></small></font></td>
    <td width="34%" align="middle" bgcolor="#FFD6D6"><font face="Verdana" color="#000000"><small><strong>Movimientos</strong></small></font></td>
  </tr>
<% IF instr(Cadena, "INFORMACION TERMINADA") > 0  THEN %>
    <tr bgcolor="#FCE8AB"><td width="67%" colspan="2" align="middle">
    <strong><small>
    <font face="Verdana">Esta cuenta no ha tenido movimientos.</font>
    </small>
    </strong>
    </td>
  </tr>
   
<% ELSE 
     while ((INSTR(CADENA, "CREDITO") > 0) or (INSTR(CADENA, "DEBITO") > 0)) 
     
       P=INSTR(CADENA, "CREDITO") 
     
       P1=INSTR(CADENA, "DEBITO")
       
       if (P1 < P) and (P1 <> 0) then P = P1 end if 
       if P = 0 then P = P1 end if
            
       P1 = INSTR(P, CADENA, ";")
         
       C1=mid(CADENA, p-10, p1-p+10)
          
       fe=mid(c1, 1, 8)
            
       If INSTR(mid(c1, 9, 9),"CREDITO") > 0 then 
         cd="Cr"
       Else
         cd="Db"
       End If
     
       mo=Trim(mid(c1, 18))
       Cadena=MID(CADENA, P1+1) 
       
     %>
     
     </td>
     </tr>
     <tr>
      <td width="33%" align="middle" bgcolor="#FCE8AB"><font size="2" face="Verdana"><%=fe%></font></td>
      <td width="34%" align="right" bgcolor="#FCE8AB"><font size="2" face="Verdana"><%=mo%><%=" "%><%=cd%></font><font color="#ffffff" size="2" face="Arial">&nbsp;</font><font color="#ffffff"><font face="Arial"><font size="2">.....</font></font><font size="2" face="Arial">.....</font></font></td>
     </tr>
   <% Wend %>
</table>

<% P2=INSTR(CADENA,"SALDO CONTABLE") %>
<% P3=INSTR(p2+14, CADENA, ";")%>
<% 
   if (P2 <> 0) and (P3 <> 0) and (P3 > P2) then 
    C2=mid(cadena, p2, p3-p2) 
   end if

   if P3 <> 0 then 
    Cadena=MID(CADENA, P3)
   end if
%>

<table width="70%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
   &nbsp;
  </tr>
  <tr>
    <td width="100%" bgcolor="#FCE8AB"><p align="right"><small><small><small><small><small><font face="Verdana"><strong>
    <% =c2 %></strong></font></small></small></small></small></small></p></td>
  </tr>
</table>

<p><% P4=INSTR(CADENA,"SALDO DISPONIBLE") %><% P5=INSTR(p4+14, CADENA, ";")%><% C3=mid(cadena, p4, p5-p4)%><% Cadena=MID(CADENA, P5)%></p>

<table width="70%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <td width="100%" bgcolor="#FCE8AB">
    <p align="right">
    <small>
    <small>
    <font face="Verdana"><strong>
    <% =c3 %></strong></font></small></small></p></td>
  </tr>
</table>

<p><% END IF%></p>
<% Else %>

<%if rs2("respcode")=91 then%>

            <table border="0" cellspacing="1" width="90%">
              <tr>
                <td width="50%" valign="center" align="middle"> <font color="#dee30f">  <IMG border=0 height=37 src="../Images/Warning.gif" width=39></font></td>
                <td width="50%" rowspan="2" valign="center" align="left"><FONT size=2><b><font face="Arial" color="#000000" >El tiempo requerido para 
      esta transacción ha sido excedido y no se ha podido establecer
          conexión con la sucursal.</font></b> </FONT>
                  <p><b><font face="Arial" color="#000000" size="2">Por favor,
                  intente de nuevo o consulte
          con el Banco.</font></b></p></td>
              </tr>
              <tr>
                <td width="50%" valign="center" align="middle"><IMG border=0 height=20 src="../Images/Tiempo%20Excedido.gif" width=182></td>
              </tr>
            </table>

<%else%>

<%  set conn = Server.CreateObject( "ADODB.Connection" )
    conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")
    query3 = "SELECT spanish FROM C_RECODE WHERE (Resp_Code = '" & rs2("respcode") & "')"
    set rs3 = conn.Execute( query3 ) %>
<%  =rs3("spanish") %>
 
<%end if%>

<% END IF%>  
<% End If%>
</body> 
</html>