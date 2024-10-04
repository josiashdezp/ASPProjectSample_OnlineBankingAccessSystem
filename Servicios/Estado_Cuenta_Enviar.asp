<% Response.Expires=0 
   Response.Buffer = True 
   
   cta = Request.Querystring("cta")
   the_date = Request.Querystring("the_date")   
   the_date1 = Request.Querystring("the_date1")
   email = Request.Form("email")   
   cc =  Request.Form("cc")   
   money = mid(cta,1,2)
   subcue = mid(cta,7,1)
   des_cuenta = mid(cta,12,2)
   
   On error resume next
   
   set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") "registro"
   
   query = "SELECT Contenido From Configuracion WHERE Variable ='SMTPHost'"
   set rsh = conn.Execute( query )
   strHost = rsh("Contenido")

   query = "SELECT Contenido From Configuracion WHERE Variable ='SMTPFrom'"
   set rsf = conn.Execute( query )
   strFrom = rsf("Contenido")
%>
<html>
<head>
<title>Envio del Estado de Cuenta por Email</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link REV="made" href="mailto:<%=strFrom%>">
<style type="text/css">
<!--
.style1 {
	font-family: Verdana;
	font-size: 12px;
	font-weight: bold;
	color: #FF0000;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #Include file = "Informa.asp"-->
<p><img border="0" src="../Images/Estado_Cuenta.gif" width="560" height="26"></p>  
<%       
    sucu = mid(Session("UsrId"), 1 , 4)
    whois = mid(Session("UsrId"), 5 , 4)
    
    set conn = Server.CreateObject( "ADODB.Connection" )
    conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")    
   
    queryb= "SELECT PLAZA FROM C_BANCOS WHERE (COD_BANCO= '" & SUCU & "')"
    set rsv = conn.Execute( queryb )
   
    histor = "HIST" & sucu 
    Cuemay = "CUE" & sucu	
     
	Set Mail = Server.CreateObject("Persits.MailSender")
	Mail.Host = strHost
	Mail.From = strFrom
	Mail.FromName = "Virtual Bandec"
	Mail.Subject = "Estado de Cuenta por Email, desde " & the_date & " hasta "& the_date1
	Mail.IsHTML = True
	Mail.Body = ""
	Mail.AddAddress email
	
	If not (cc="") then
	  Mail.AddCc cc
	End If
	
	mensaje="<html><head>"
    mensaje=mensaje & " <title>Estado de Cuenta por Email</title> "&vbcrlf
    mensaje=mensaje & " </head><body><div align=center> "&vbcrlf
	mensaje=mensaje & " <P><B><div align=left><font color=#000000 face=Verdana size=3>Estado de Cuenta.</font></b></p>"&vbcrlf
	mensaje=mensaje & " <hr size=1 color=#000000 align=left> "&vbcrlf
	mensaje=mensaje & " <div align=right><font color=#000000 size=3 face=Verdana> "&vbcrlf
	mensaje=mensaje & " <pre><font color=#000000 face=Verdana size=2>BANDEC, SUCURSAL " & sucu & ", " & rsv("PLAZA") & "</font>"&vbcrlf
	mensaje=mensaje & " <font color=#000000 face=Verdana size=2>Fecha Emisión: " & Date() & " </font>"&vbcrlf
	mensaje=mensaje & " <font color=#000000 face=Verdana size=2>Cuenta: " & cta & " </strong></pre>"&vbcrlf
	mensaje=mensaje & " <div align=center><center><table border=0 cellpadding=0 cellspacing=5 width=800><tr><td align=center colspan=6><hr color=#000000 size=1></td></tr>"&vbcrlf
	mensaje=mensaje & " <tr><td width=8% align=center><p align=center><font face=Arial size=1><strong>Fecha Contable</strong></span></font></td>"&vbcrlf
    mensaje=mensaje & " <td width=10% align=center><p align=center><font face=Arial size=1><strong>Referencia Corriente</strong></span></font></td>"&vbcrlf
	mensaje=mensaje & " <td width=10% align=center><p align=center><font face=Arial size=1><strong>Referencia Original</strong></span></font></td>"&vbcrlf
	mensaje=mensaje & " <td width=50% align=center><font face=Arial size=1><strong>Observaciones&nbsp;&nbsp;</strong></span></font></td><td width=10 align=center></td><td width=19 align=center><font face=Arial size=1><span style=text-transform: uppercase><strong>Movimientos</strong></span></font></td></tr>"&vbcrlf
	mensaje=mensaje & " <tr><td colspan=6><hr color=#000000 size=1></td></tr><tr><td width=12></td><td width=16></td><td width=15></td><td colspan=2><p align=right><strong><span style=text-transform: uppercase><font face=Arial size=1>Saldo Anterior: </font></span></strong></td>"&vbcrlf

  If mid(right(Cta,14),4,4) = "3360" then
    money = mid(right(Cta,14),1,3)
    SUBCUE = "3360"
	des_cuenta = mid(right(cta,14),13,2)
  Else
  If mid(right(cta,14),1,2) = "40" then
    money = "CUP"
   Else
  If mid(right(cta,14),1,2) = "43" then
    money = "CUC"
   Else   
    money = "USD"
   End if
   End if
   des_cuenta = mid(right(cta,14),12,2)
   If mid(right(cta,14),7,1) = "1" then
    SUBCUE = "3210"
   Else
    If mid(right(cta,14),7,1) = "8" then
	 SUBCUE = "3280"
	Else
	 SUBCUE = "3290"
	End If
   End If
  End If 
	
    querysa = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & Des_cuenta & "') AND (FEC_CONTAB >= {d '" & the_date & "'}) AND (COD_ASIENT='120') Order by Fec_Contab"	
    set rssa = conn.Execute( querysa )
    
	If rssa("IMP_ASIENT") < 0 then cod = "Cr" else cod = "Db" End If
	Imp_asient = FormatNumber(Abs(rssa("IMP_ASIENT")),2) & cod
	mensaje=mensaje & " <td align=right><font face=Arial><font size=2>" & Imp_asient & " </font></font></td></tr> "&vbcrlf
	mensaje=mensaje & " <tr><td width=12></td><td width=16></td><td width=15></td><td width=28><font size=1></font></td><td colspan=2><font size=1><hr color=#000000 width=100 size=1 align=right></font></td></tr></table></center></div>"&vbcrlf
	
    query1 = "SELECT FEC_CONTAB, REF_CORRIE, REF_ORIGIN, OBSERV, IMP_ASIENT, COD_ASIENT FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB>={d '" & the_date & "'}) AND (FEC_CONTAB<={d '" & the_date1 & "'}) AND ((IsNull(COD_ASIENT))  OR ((not IsNull(COD_ASIENT)) AND (COD_ASIENT <> '120') AND (COD_ASIENT <> '121') AND (COD_ASIENT <> '122') AND (COD_ASIENT <> '123') AND (COD_ASIENT <> '125') AND (COD_ASIENT <> '126'))) Order by Fec_Contab ASC"	
    set rs1 = conn.Execute( query1 )  
    
	Do While not rs1.eof 
    
	If rs1("COD_ASIENT") <> "124" OR IsNull(rs1("COD_ASIENT")) then
    mensaje=mensaje & " <div align=center> "&vbcrlf 

	mensaje=mensaje & " <div align=center><center> "&vbcrlf
	mensaje=mensaje & " <table border=0 cellpadding=0 cellspacing=5 width=800><tr> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><p align=center><font face=Arial size=1> " & rs1("FEC_CONTAB")& " </font></td> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><font face=Arial><font size=1> " & rs1("REF_CORRIE")& " </font></font></td> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><p align=center><font face=Arial size=1> " & rs1("REF_ORIGIN") & " </font></td> "&vbcrlf
	mensaje=mensaje & " <td width=50% align=center><p align=right><font face=Arial size=1> " & rs1("Observ") & " </font></td><td width=3></td></center> "&vbcrlf
    
	If rs1("IMP_ASIENT") < 0 then cod = "Cr" else cod = "Db" End If
	Imp_asient = FormatNumber(abs(rs1("IMP_ASIENT")),2) & cod 
    mensaje=mensaje & " <td align=right valign=Top><font size=2><font face=Arial>" & Imp_Asient & "&nbsp;&nbsp;</font></font></td></tr></table></div></div> "&vbcrlf
	Else
	mensaje=mensaje & " <div align=center><center><table border=0 cellpadding=0 cellspacing=5 width=800><tr> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><small><small><font face=Arial> " & rs1("FEC_CONTAB") & " </font></small></small></td> "&vbcrlf
	mensaje=mensaje & " <td align=left><strong><small><small><font face=Arial> No hubo movimientos en esta fecha. </font></small></small></strong></td></tr></table></center></div> "&vbcrlf
	End If
    rs1.MoveNext 
    Loop 
	
	mensaje=mensaje & " <div align=center> "&vbcrlf
    queryfecha = "SELECT FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB >= {d '" & The_date & "'}) AND (FEC_CONTAB <= {d '" & The_date1 & "'})AND (COD_ASIENT='121')"
    set rsfecha = conn.Execute( queryfecha ) 
    fecult=# 1-1-1900 #
    Do While not rsfecha.eof
    fec = FormatDateTime(rsfecha("Fec_Contab"))
    If fec > fecult then
      fecult=fec  
      fechault=datepart("yyyy", fecult ) & "-" & mid(rsfecha("FEC_CONTAB"), 4, 2) & "-" & mid(rsfecha("FEC_CONTAB"), 1, 2)
    End If 
    rsfecha.movenext 
    LOOP 

	mensaje=mensaje & " <div align=center><center> "&vbcrlf
	mensaje=mensaje & " <table border=0 cellpadding=0 cellspacing=0 width=800><tr><td width=9 rowspan=2></td><td width=9 rowspan=2></td><td width=9 rowspan=2></td><td width=9 rowspan=2></td><td colspan=2 rowspan=2><font size=1></font></td><td rowspan=2 width=18><font size=1><hr color=#000000 size=1 width=100 align=right></font></td></tr> "&vbcrlf
	mensaje=mensaje & " <tr></tr><tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size=1><span style=text-transform: uppercase><b>Saldo final:</b></span></font></td>"&vbcrlf
	
	queryscont = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='121') Order By Fec_Contab DESC"
    set rsscont = conn.Execute( queryscont )
    
	If rsscont("IMP_ASIENT") < 0 then cod = "Cr" else cod = "Db" End If
	Imp_Asient = FormatNumber(Abs(rsscont("IMP_ASIENT")),2) & cod
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	
    if (subcue = "3280") or (subcue = "3290") then Saldo = "Fondo aprobado:" else  Saldo = "Sobregiro autorizado:" end if
	mensaje=mensaje & " <tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size=1><span style=text-transform: uppercase><b>" & Saldo & "</b></span></font></td>"&vbcrlf

    querysf = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='126') Order By Fec_Contab DESC"
    set rssf = conn.Execute( querysf )
	If rssf.Eof then Impt=0 else Impt=rssf("IMP_ASIENT") end if
	If IMPT < 0 then cod = "Cr" else cod = "Db" End If
    Imp_Asient = FormatNumber(Abs(IMPT),2) & cod
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	
	mensaje=mensaje & " <tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size=1><span style=text-transform: uppercase><b>Fondo Reservado: </b></span></font></td>"&vbcrlf
	
	querysdisp = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='125') Order By Fec_Contab DESC"
    set rssdisp = conn.Execute( querysdisp )
    
	If rssdisp("IMP_ASIENT") < 0 then cod = "Cr" else cod = "Db" End If
	Imp_Asient = FormatNumber(Abs(rssdisp("IMP_ASIENT")),2) & cod
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	
	mensaje=mensaje & " <tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size=1><span style=text-transform: uppercase><b>Fondo Disponible:</b></span></font></td>"&vbcrlf

	querysdisp = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='123') Order By Fec_Contab DESC"
    set rssdisp = conn.Execute( querysdisp )
	If rssdisp("IMP_ASIENT") < 0 then cod = "Cr" else cod = "Db" End If
	Imp_Asient = FormatNumber(Abs(rssdisp("IMP_ASIENT")),2) & cod
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	
	mensaje=mensaje & " </table></center></div></div> "&vbcrlf
	mensaje=mensaje & " <hr size=1 color=#000000 align=left width=800>"&vbcrlf
	mensaje=mensaje & " <font face=verdana size=1><div align=center><b>Un servicio de Virtual Bandec.</b></div></font> "&vbcrlf
	Mail.Body = mensaje
	Mail.Send
	
	If Err <> 0 Then %>
     <center><IMG SRC="../Images/AccesoDenegado.gif" WIDTH="45" HEIGHT="45" BORDER=0 ALT="Operación Satisfactoria"></center>
<div align="center"><BR>
    <font  color="#3300FF" face="Verdana, Arial, Helvetica, sans-serif"> <B><span class="style1">Ha ocurrido un error. Es posible que el Estado de Cuentas NO se haya enviado correctamente.</span>
    </div>
     <p>
	 <% Response.Write "Descripción del error: " & Err.Description 
        else %>
     <center><IMG SRC="../Images/PasswordExito.gif" WIDTH="45" HEIGHT="45" BORDER=0 ALT="Operación Satisfactoria"></center><BR>
     <center><font face="verdana" size="2"><B><%Response.Write "<font face=verdana size=2 color=#3300FF>El Estado de Cuenta fue enviado satisfactoriamente a la siguiente dirección:&nbsp;</font>"&email
	End If %>
	</B></font></center>
</body>
</html>