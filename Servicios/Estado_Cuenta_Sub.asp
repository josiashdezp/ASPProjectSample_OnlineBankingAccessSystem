<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>Estado de Cuenta</title>
</head>
<script Language="VBScript">
  Sub Hand()
    Alert "El año 2000 es bisiesto"
  end Sub
</script>
 <style>
 
 body { scrollbar-arrow-color: #DDCE67;
scrollbar-base-color: #f0b6c4;
scrollbar-face-color: #9f314e;
scrollbar-highlight-color: #f0b6c4;
scrollbar-shadow-color: #DDCE67; }

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
	color: #000000;
	font-size: 18px;
	filter:glow(color=#FFFFFF,strength=1);
width:100%;
}
.style3 {
	font-family: Verdana;
	color: #FFFFCC;
	font-size: 14px;
	filter:glow(color=#000000,strength=2);
width:100%;
}
.style4 {
	font-family: Verdana;
	color: #000000;
	font-size: 12px;
	
}
.style5 {
	font-family: Verdana;
	color: #000000;
	font-size: 10px;
	
}
</style>
<!-- #Include file = "Informa.asp"-->
<body style="background-color:transparent; ">
 <table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style1">Estado de Cuentas.</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
  
</table>
<form METHOD="POST" ACTION="Estado_Cuenta_Sub.asp" id=form1 name=form1>
  <input type="hidden" name="hname" value="hvalue">
<% 
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4) 

   histor = "HIST" & sucu 
   Cuemay = "CUE" & sucu

 If Request.Form("hname") = "" Then 

   Function Resto (cuenta)
     Cuenta=TRIM(Cuenta)
     SUMA=0

     DIM PESO(13)
     PESO(1)=1
     PESO(2)=2
     PESO(3)=3
     PESO(4)=5
     PESO(5)=7 
     PESO(6)=11
     PESO(7)=13
     PESO(8)=17
     PESO(9)=19
     PESO(10)=21 
     PESO(11)=23
     PESO(12)=29
     PESO(13)=31
     FOR I=1 TO LEN(Cuenta)
       SUMA=SUMA+(CINT(MID(Cuenta,I,1)) * PESO(I))
     NEXT 
     RESTO = TRIM (CSTR(SUMA MOD 11))  
     IF RESTO <> 10 then 
      RESTO = RESTO
     Else 
      RESTO = MID(Cuenta,7,1)
     END IF           
   End Function

  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")  %>
  <table border="0" width="98%" cellspacing="8" cellpadding="8" align="center">
    <tr>
      <td align="right" valign="middle" bgcolor="#FFD6D6">
        <div align="center"><span><strong><small><font face="Verdana" size="2">Seleccione una Cuenta:</font></small></strong></span>
            <select name="Cta_Debito" size="1">
        <option value="Todas">Todas las cuentas</option>
      
    <% 'Esto es para listar las cuentas de los clientes subordinados al que esta en el sistema realmente  
  
 query = "SELECT * FROM subclien WHERE (CUE_SUCUR = '" & sucu & "') and (COD_CONTRA = '" & whois & "')" 
   set subc = conn.Execute(query) 
    do while not subc.EOF     
     file = "CUE" & subc("sucu") 
     query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM "& file &" WHERE (Cod_Contra = '" & subc("whois") & "') AND (Cue_Inact = .F.) AND (Cue_Cierre = .F.) AND (Cue_Cierre = .F.)"
     set rs = conn.Execute( query ) 
      do while not rs.EOF        
      if (rs("cue_subcue") = "3360") or (rs("cue_subcue") = "3210") or (rs("cue_subcue") = "3280") or (rs("cue_subcue") = "3290") then 
      if (rs("cue_subcue") = "3360") then %>
	          <option value="
        <% cta = rs("sig_moneda") & rs("cue_subcue") & rs("tip_contra") & subc("whois") & rs("des_cuenta") %>   
        <%=cta%>">
              <%=cta%> </option>
      
          <%else   
         query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
	     set rs2 = conn.Execute( query2 ) 
	     money = rs2("Cod_Moneda") 
    
		cta = money & subc("sucu") & mid(rs("cue_subcue"), 3, 1) & subc("whois") & rs("des_cuenta") %>   	
              <option value="<%=cta%><%=Resto(cta)%>"><%=cta%><%=Resto(cta)%> </option>
    	      <% end if 
	  end if    
   rs.MoveNext 
   loop
   subc.MoveNext
   loop%> 
            </select>     
            <input TYPE="submit" VALUE="BUSCAR" style="color:#FFFFFF;font-family: Verdana; font-weight: bold;background-color:#9f314e;border-style:outset" id="submit1" name="submit1">
        </div></td>
    </tr>
    <tr>
      <td valign="top" align="left" bgcolor="#FCE8AB"><strong><font face="Verdana" color="#000000" size="4">M</font><font face="Verdana" color="#000000" size="2">ediante este servicio usted puede conocer el estado de una o de todas
      sus cuentas.</font></strong></td>
    </tr>
  </table>
  </center></div>
</form>         
<%else
 MostrarEmail = false
 Cta_Debito = cstr(Request.Form("Cta_Debito")) %>
<span class="style2">Seleccione el rango de fecha deseado:</span>

<%  function MonthName( month )
    dim month_names(12)
    month_names(  1 ) = "Enero"
    month_names(  2 ) = "Febrero"
    month_names(  3 ) = "Marzo"
    month_names(  4 ) = "Abril"
    month_names(  5 ) = "Mayo"
    month_names(  6 ) = "Junio"
    month_names(  7 ) = "Julio"
    month_names(  8 ) = "Agosto"
    month_names(  9 ) = "Septiembre"
    month_names( 10 ) = "Octubre"
    month_names( 11 ) = "Noviembre"
    month_names( 12 ) = "Diciembre"
    MonthName = month_names(month)
  end function

  Function Formalize( date_part )
    if CInt( date_part ) < 10 then
      Formalize = "0" & date_part
    else
      Formalize = date_part
    end if
   End function

  yy = Request.Form( "year"  )
  mm = Request.Form( "month" )
  dd = Request.Form( "day"   )

  if (yy="") or (mm="") or (dd="") then
    yy = DatePart( "yyyy", Date() )
    mm = Formalize( DatePart( "m", Date() ) )
    dd = Formalize( DatePart( "d", Date() ) )
  end if

  the_date = yy & "-" & mm & "-" & dd

  yy1 = Request.Form( "year1"  )
  mm1 = Request.Form( "month1" )
  dd1 = Request.Form( "day1"   )

  if (yy1="") or (mm1="") or (dd1="") then
    yy1 = DatePart( "yyyy", Date() )
    mm1 = Formalize( DatePart( "m", Date() ) )
    dd1 = Formalize( DatePart( "d", Date() ) )
  end if

  the_date1 = yy1 & "-" & mm1 & "-" & dd1

%>
<form action="Estado_Cuenta_Sub.asp" method="POST">
 <table width="100%" border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td bgcolor="#FCE8AB"><p><strong><font face="Arial" size="2">Desde:</font></strong>
	    <input type="hidden" name="hname" value="hvalue">
      <input type="hidden" name="Cta_Debito" value="<%=Cta_Debito%>"></td>
    <td bgcolor="#FCE8AB"><font face="Arial" size="2"><strong>Hasta</strong></font>:</td>
  </tr>
  <tr>
    <td bgcolor="#FCE8AB">
	      <select name="day" size="1">
<% for d = 1 to 31 %><% if d = CInt(dd) then %> 
       <option selected value="<%=Formalize(d)%>"><%=d%></option>
<% else %>
        <option value="<%=Formalize(d)%>"><%=d%></option>
<% end if %>
<% next %>
      </select>
	  <select name="month" size="1">
<% for m = 1 to 12 %><% if m = CInt(mm) then %>
        <option selected value="<%=Formalize(m)%>"><%=MonthName(m)%></option>
<% else %>
        <option value="<%=Formalize(m)%>"><%=MonthName(m)%></option>
<% end if %>
<% next %>
      </select>
	  <select name="year" size="1">
<% for y = 1998 to Year(Date()) %><% if y = CInt(yy) then %>        <option selected value="<%=yy%>"><%=yy%></option>
<% else %>        <option value="<%=y%>"><%=y%></option>
<% end if %><% next %>      </select>	</td>
    <td bgcolor="#FCE8AB"> <select name="day1"     size="1">
<% for d = 1 to 31 %><% if d = CInt(dd1) then %>        <option selected value="<%=Formalize(d)%>"><%=d%></option>
<% else %>        <option value="<%=Formalize(d)%>"><%=d%></option>
<% end if %><% next %>      </select>
<select name="month1" size="1">
<% for m = 1 to 12 %><% if m = CInt(mm1) then %>        <option selected value="<%=Formalize(m)%>"><%=MonthName(m)%></option>
<% else %>        <option value="<%=Formalize(m)%>"><%=MonthName(m)%></option>
<% end if %><% next %>      </select>
<select name="year1" size="1">
<% for y = 1998 to Year(Date()) %><% if y = CInt(yy1) then %>        <option selected value="<%=yy1%>"><%=yy1%></option>
<% else %>        <option value="<%=y%>"><%=y%></option>
<% end if %><% next %>      </select>&nbsp;&nbsp;<input type="submit" name="go" value="Buscar" width="2"></td>
  </tr>
  <tr>
    <td colspan="2">
</td>
  </tr>
</table>
</form>

<% If the_date > the_date1 or Month(the_date) > Month(Date()) and Year(the_date) >= Year(Date()) or Day(the_date) > Day(Date()) and Month(the_date) >= Month(Date()) and Year(the_date) >= Year(Date()) or Year(the_date) > Year(Date()) or Month(the_date1) > Month(Date()) and Year(the_date1) >= Year(Date()) or Day(the_date1) > Day(Date()) and Month(the_date1) >= Month(Date()) and Year(the_date1) >= Year(Date()) or Year(the_date1) > Year(Date()) then%>
<span class="style2"><% Response.write("Rango de Fecha no válido. Verifique el Rango de Fechas.")%></span>
<% MostrarEmail = false
Else
  MostrarEmail = true
  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 

   Function Resto (cuenta)
     Cuenta=TRIM(Cuenta)
     SUMA=0

     DIM PESO(13)
     PESO(1)=1
     PESO(2)=2
     PESO(3)=3
     PESO(4)=5
     PESO(5)=7 
     PESO(6)=11
     PESO(7)=13
     PESO(8)=17
     PESO(9)=19
     PESO(10)=21 
     PESO(11)=23
     PESO(12)=29
     PESO(13)=31
     FOR I=1 TO LEN(Cuenta)
       SUMA=SUMA+(CINT(MID(Cuenta,I,1)) * PESO(I))
     NEXT 
     RESTO = TRIM (CSTR(SUMA MOD 11))  
     IF RESTO <> 10 then 
      RESTO = RESTO
     Else 
      RESTO = MID(Cuenta,7,1)
     END IF           
   End Function


 query = "SELECT * FROM subclien WHERE (CUE_SUCUR = '" & sucu & "') and (COD_CONTRA = '" & whois & "')" 
   set subc = conn.Execute(query) 

 do while not subc.EOF     

file = "CUE" & subc("sucu") 

   query = "SELECT * FROM " & file & " WHERE (COD_CONTRA='" & subc("whois") & "')"
     set rs = conn.Execute( query ) 

 Do While Not rs.EOF 

 Mostrar = False 

if Cta_Debito = "Todas" then
  If rs("CUE_SUBCUE") = 3210 Or rs("CUE_SUBCUE") = 3290 Or rs("CUE_SUBCUE") = 3280 Or rs("CUE_SUBCUE") = 3360 Then
    Mostrar = True
  end if
else
  if (mid(rs("cue_subcue"),1,3) <> "141") and (mid(rs("cue_subcue"),1,3) <> "151") then

  query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
  set rs2 = conn.Execute( query2 ) 
  money = rs2("Cod_Moneda") 

  if rs("CUE_SUBCUE") = 3360 then
    cta = rs("sig_moneda") & rs("cue_subcue") & rs("tip_contra") & subc("whois") & rs("des_cuenta")
  else
    cta = money & subc("sucu") & mid(rs("cue_subcue"), 3, 1) & subc("whois") & rs("des_cuenta")
    cta = cta & Resto(cta)
  end if 
  if instr( Cta_Debito, cta ) > 0 then
    Mostrar = True
  end if
  
  end if
end if 

 If Mostrar = True Then %>

<hr color="#FFFFCC" noshade size="1">
<%queryb= "SELECT PLAZA FROM C_BANCOS WHERE (COD_BANCO= '" & subc("sucu") & "')"%>
<% set rsb = conn.Execute(queryb)%>
<table width="100%%"  border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td><span class="style2">BANDEC, Sucursal <% = subc("sucu") %>, <%=rsb("Plaza")%></span></td>
  </tr>
  <tr>
    <td><span class="style2">Fecha Emisión: <%Response.Write(Date())%></span></td>
  </tr>
    <tr>
    <td><span class="style2">Cuenta: <% if rs("cue_subcue") = 3360 then %>
  <% cta = rs("sig_moneda") & rs("cue_subcue") & rs("tip_contra") & subc("whois") & rs("des_cuenta") %>   
  <%=cta%>
<% else 
 
  query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
  set rs2 = conn.Execute( query2 ) 
  money = rs2("Cod_Moneda")
    
  cta = money & subc("sucu") & mid(rs("cue_subcue"), 3, 1) & subc("whois") & rs("des_cuenta") %>
  <%=cta%><%=Resto(cta)%>
<% end if %></span>
</td>
  </tr>
</table>

    <% histor = "HIST" & subc("sucu") 
       query1 = "SELECT FEC_CONTAB, REF_CORRIE, REF_ORIGIN, OBSERV, IMP_ASIENT, COD_ASIENT FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB>={d '" & the_date & "'}) AND (FEC_CONTAB<={d '" & the_date1 & "'}) AND ((IsNull(COD_ASIENT))  OR ((not IsNull(COD_ASIENT)) AND (COD_ASIENT <> '120') AND (COD_ASIENT <> '121') AND (COD_ASIENT <> '122') AND (COD_ASIENT <> '123') AND (COD_ASIENT <> '125') AND (COD_ASIENT <> '126'))) Order by Fec_Contab ASC"
    set rs1 = conn.Execute( query1 ) 
    If rs1.eof then %>
   
    <span style="background-color:#FCE8AB" class="style2">No existe información para este rango de fecha. Intente más tarde o consulte con el Banco.</span>
<% Else %>

<table border="0" cellpadding="3" cellspacing="1" width="100%">

   <tr bgcolor="#FCE8AB">
    <td width="11%" align="center"><span class="style5">Fecha Contable</span></td>
    <td width="11%" align="center"><span class="style5">Refer. Corriente</span></td>
    <td width="11%" align="center"><span class="style5">Refer. Original</span></td>
    <td width="50%" align="center" bgcolor="#FCE8AB"><span class="style5">Observaciones</span></td>
    <td width="17%" align="center"><span class="style5">Movimientos</span></td>
  </tr>
  <tr>
    <td width="100%" colspan="5"><hr size="1" color="#000000">
    </td>
  </tr>
  <tr>
    <td width="11%"></td>
    <td width="11%"></td>
    <td width="11%"></td>
    <td width="50%" align="right" bgcolor="#FCE8AB"><span class="style5">Saldo Anterior:</span></td>
    <td width="17%" align="right" bgcolor="#FCE8AB"><span class="style5">
	<%
   querysa = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB >= {d '" & the_date & "'}) AND (COD_ASIENT='120') Order by Fec_Contab"
   set rssa = conn.Execute( querysa )
   If CDbl(rssa("IMP_ASIENT")) < 0 then cod = " Cr" else cod = " Db" End If%>
   <%=FormatNumber(Abs(CDbl(rssa("IMP_ASIENT"))),2)%>
   <%=cod%><font color="#FFFFFF"></span></td>
  </tr>
 
</table>
<% Do While not rs1.eof 
 If rs1("COD_ASIENT") <> "124" OR IsNull(rs1("COD_ASIENT")) then%>

<% Obs = rs1("observ")  
   Longitud= Len(Obs)
   
 Do while inStr(OBS,";")<>0 
 Posicion = inStr(Obs,";")
 Obs = Mid(Obs,1,Posicion-1) & "  " & Mid(Obs,Posicion+1) 
 Loop 
 
  If inStr(Obs,"Comprobante:")<>0 then 
   Posicion = inStr(Obs,"Comprobante:")
   Obs = Mid(Obs,1,Posicion-1) & "  " & Mid(Obs,Posicion) 
   Posicion = inStr(Obs,"REF_CORRIE:")
   Obs = Mid(Obs,1,Posicion-1) & "  " & Mid(Obs,Posicion) 
   Posicion = inStr(Obs,"FEC_CONTAB:")
   Obs = Mid(Obs,1,Posicion-1) & "  " & Mid(Obs,Posicion)
 end if %>
<table border="0" cellpadding="3" cellspacing="1" width="100%"> 
  <tr bgcolor="#FFCCCC">
    <td width="11%" align="center" valign="Top"><p align="center"><small><small><font
    face="Arial"><% =rs1("FEC_CONTAB") %></font></small></small></td>
    <td width="11%" align="center" valign="Top"><small><small><font face="Arial"><% =rs1("REF_CORRIE") %></font></small></small></td>
    <td width="11%" align="center" valign="Top"><p align="center"><small><small><font
    face="Arial"><% =rs1("REF_ORIGIN") %></font></small></small></td>
    <td width="50%" align="center"><p align="right"><small><small><font face="Arial"><% = Obs %> </font></small></small></td>
    <td width="17%" align="right" valign="top"><small><small><font face="Arial">
      <%If CDbl(rs1("IMP_ASIENT")) < 0 then cod = " Cr" else cod = " Db" End If%> 
    </font><small><font face="Arial">
    <% =FormatNumber(abs(CDbl(rs1("IMP_ASIENT"))),2)&cod %></font></small></small></small></td>
  </tr>
</table>
<%else%>
<table border="0" cellpadding="3" cellspacing="1" width="100%">
 <tr bgcolor="#FFCCCC">
    <td width="11%" align="center"><small><small><font face="Arial"><% =rs1("FEC_CONTAB") %></font></small></small></td>
    <td width="11%"></td>
	<td width="11%"></td>
    <td width="50%" align="left"><strong><small><small><font face="Arial">No hubo movimientos
    en esta&nbsp; fecha.</font></small></small></strong></td>
	<td width="17%"></td>
  </tr>
</table>
<%end if
 rs1.MoveNext 
 Loop %>
<% queryfecha = "SELECT FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB >= {d '" & The_date & "'}) AND (FEC_CONTAB <= {d '" & The_date1 & "'})AND (COD_ASIENT='121')"
   set rsfecha = conn.Execute( queryfecha )%>
<% fecult=# 1-1-1900 # 
Do While not rsfecha.eof
 fec = FormatDateTime(rsfecha("Fec_Contab"))
 If fec > fecult then
      fecult=fec  
      fechault=datepart("yyyy", fecult ) & "-" & mid(rsfecha("FEC_CONTAB"), 4, 2) & "-" & mid(rsfecha("FEC_CONTAB"), 1, 2)
   End If 
 rsfecha.movenext 
 LOOP %>
<table border="0" cellpadding="3" cellspacing="1" width="100%" ID="Table1">
  <tr>
    <td width="11"></td>
    <td width="11%"></td>
    <td width="11%"></td>
    <td width="50%" align="right" bgcolor="#FCE8AB"><span class="style5">Saldo final:</span></td>
    <td width="17%" align="right" bgcolor="#FCE8AB"><span class="style5">
      <% queryscont = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='121') Order By Fec_Contab DESC"
       set rsscont = conn.Execute( queryscont )%>
       <%If CDbl(rsscont("IMP_ASIENT")) < 0 then cod = " Cr" else cod = " Db" End If%>
       <%=FormatNumber(Abs(CDbl(rsscont("IMP_ASIENT"))),2)%><%=cod%></span></td>
  </tr>
  <tr>
    <td width="11"></td>
    <td width="11%"></td>
    <td width="11%"></td>
    <td width="50%" align="right" bgcolor="#FCE8AB"><span class="style5"><%if (rs("CUE_SUBCUE") = "3280") or (rs("CUE_SUBCUE") = "3290") then%>Fondo aprobado:<%else%>Sobregiro autorizado: </span>
      <%end if%>
    </td>
    <td width="17%" align="right" bgcolor="#FCE8AB"><span class="style5">
       <% querysconf = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='126') Order By Fec_Contab DESC"
       set rssconf = conn.Execute( querysconf )
	   If CDbl(rssconf("IMP_ASIENT")) < 0 then cod = " Cr" else cod = " Db" End If
       If rssconf.Eof then Impt=0 else Impt=CDbl(rssconf("IMP_ASIENT")) end if%>
	   <%=FormatNumber(Abs(Impt),2)%>
	   <%=cod%></span></td>
  </tr>
  <tr>
    <td width="11"></td>
    <td width="11%"></td>
    <td width="11%"></td>
    <td width="50%" align="right" bgcolor="#FCE8AB"><span class="style5">Fondo reservado:</span></td>
    <td width="17%" align="right" bgcolor="#FCE8AB"><span class="style5">
       <% querysdisp = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='125') Order By Fec_Contab DESC"
       set rssdisp = conn.Execute( querysdisp )
       If CDbl(CDbl(rssdisp("IMP_ASIENT"))) < 0 then cod = " Cr" else cod = " Db" End If%>
       <%=FormatNumber(Abs(CDbl(CDbl(rssdisp("IMP_ASIENT")))),2)%><%=cod%></span></td>
  </tr>
  <tr>
    <td width="11"></td>
    <td width="11%"></td>
    <td width="11%"></td>
    <td width="50%" align="right" bgcolor="#FCE8AB"><span class="style5">Fondo disponible:</span></td>
    <td width="17%" align="right" bgcolor="#FCE8AB"><span class="style5">
       <% querysdisp = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & rs("SIG_MONEDA") & "') AND (CUE_SUBCUE='" & rs("CUE_SUBCUE") & "') AND (COD_CONTRA='" & subc("whois") & "') AND (DES_CUENTA='" & rs("DES_CUENTA") & "') AND (FEC_CONTAB <= {d '" & The_Date1 & "'}) AND (COD_ASIENT='123') Order By Fec_Contab DESC"
       set rssdisp = conn.Execute( querysdisp )
       If CDbl(rssdisp("IMP_ASIENT")) < 0 then cod = " Cr" else cod = " Db" End If %>
       <%=FormatNumber(Abs(CDbl(rssdisp("IMP_ASIENT"))),2)%><%=cod%></span></td>
  </tr>
</table>
<% End If %>
<!-- No existe Información para esta fecha -->
<% End If %>
<!-- Validación de la Cuenta -->
<% rs.MoveNext 
 Loop 
 subc.MoveNext 
 Loop 
 End If %>
<!-- Validación de la fecha -->
<hr color="#FCE8AB" noshade size="1">
</div>
<%end if%>
<!-- Acá va el formulario para mandar el mensaje por correo. -->
<% if MostrarEmail then %>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="2" id="B2" >
  <tr> 
    <td colspan="2" align="left" bgcolor="#FCE8AB" style="height: 12px"><span class="style6">Ahora Ud. podrá enviar este Estado de Cuenta por Correo Electrónico (email).</span></td>
  </tr>
  <tr> 
    <td width="100%" align="center">
	  <form name="form1" method="post" action="../Lib/CrearMensaje.asp?the_date=<%=the_date%>&the_date1=<%=the_date1%>&cta=<%=right(Cta_Debito,14)%>&SubCuenta=true">
        <table width="100%" border="0" align="center" id="TablaCorreo">
        <tr>
        </tr>        
        <tr>      
          <td width=20% align="right"><font face="verdana" size="2"><b>Para:</b></font></td>
          <td align="left"><input type="text" name="email" size="20">&nbsp; <i>A la persona que desee enviarlo</i></td>
        </tr>
        <tr>
          <td width=20% align="right"><font face="verdana" size="2"><b>Con Copia a:</b></font></td>
          <td align="left"><input type="text" name="cc" size="20">&nbsp; <i>Si quiere enviar copia del mismo</i></td>        
        </tr>
            <tr align="left">
                <td align="right" colspan="2" style="text-align: center">
        <input type="submit" value="Enviar" name="B1" style="color: #FFFFFF; background-color: #271785; font-family: Verdana; font-weight: bold; border-style: outset"></td>
            </tr>
            <tr align="left" bgcolor="#FCE8AB">
                <td align="right" colspan="2" style="text-align: left" >
                    Guardar estado de Cuentas a Disco:&nbsp; Para esta operación inserte primeramente
                    el disco de 3 1/2 Pulgadas en la torre, luego <a href="#" onclick="window.open 'Estado_Cuentas_Salvar.asp?the_date=<%=the_date%>&the_date1=<%=the_date1%>&cta=<%=right(Cta_Debito,14)%>&SubCuenta=true','SubMenu','height=300,width=400,resizable,scrollbars,statusbar' " style="font-weight:bold;">Guarde Este Documento</a></td>
            </tr>
        <tr>
        </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<%end if %>
</body>
</html>