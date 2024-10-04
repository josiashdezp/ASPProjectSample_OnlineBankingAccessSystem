<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Ultimos 10 Movimientos de la Cuenta</title>
<style>
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
</style>
</head>
<!-- #Include file ="Informa.asp"-->
<body leftmargin="0" style="background-color: transparent ">
<table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style1">Ultimos 10 Movimientos.</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
  </tr>
   <tr>  
    <td width="100%" align="center"><span class="style2">Fecha: <%=Date%> - <%=Time%></span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
  </tr>
</table>
<% 
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4) 

   Application.Lock 
   Application("NumVisits") = Application("NumVisits") + 1 
   Application.Unlock 

  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")  
  
  'On Error Resume Next
   If Request.Form("hname") = "" Then
   ' Permite la captación de datos en la forma HTML

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

  file = "CUE" & sucu 
  query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM "& file &" WHERE (Cod_Contra = '" & whois & "') AND (Cue_Inact = .F.) AND (Cue_Cierre = .F.) AND (Cue_Cierre = .F.) AND (CUE_SUBCUE <> '3360')  AND ( SUBSTR(Cue_SUBCUE,1,2)  <> '14') AND (SUBSTR(Cue_SUBCUE,1,2)  <> '15')"
  set rs = conn.Execute( query )
%>
<form METHOD="POST" ACTION="10_Mov.asp">
  <input type="hidden" name="hname" value="hvalue"><div align="left"><table width="98%" border="0" cellpadding="10" cellspacing="10">
    <tr bgcolor="#FFD6D6">
      <td width="40%" align="right" valign="middle">
        <p align="right"><strong><small><font face="Verdana">Seleccione una Cuenta:</font></small></strong></p>      </td>
      <td width="30%" align="center" valign="middle" bgcolor="#FFD6D6">
        <p align="left"><select name="Cta_Debito" size="1">
<% do while not rs.EOF  
   if (mid(rs("cue_subcue"),1,3) <> "141") and (mid(rs("cue_subcue"),1,3) <> "151") then 
      query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
      set rs2 = conn.Execute( query2 ) 
      if not rs2.eof then
        money = rs2("Cod_Moneda") 
      else
       money = "??"
      end if
      cta = money & sucu & mid(rs("cue_subcue"), 3, 1) & whois & rs("des_cuenta") %>
    <option value="<%=cta%><%=Resto(cta)%>"><%=cta%><%=Resto(cta)%></option>
<% end if    
 rs.MoveNext 
  loop %>      </select></p>      </td>
       <td width="30%" align="center" valign="middle"><input TYPE="submit" VALUE="Buscar" style="color: #FFFFFF; font-family: Verdana; font-weight: bold; background-color: #9D2C4A; border-style: outset"></td>
    </tr>
    <tr bgcolor="#FCE8AB">
      <td colspan="3" align="left" valign="top"><strong><font face="Verdana" color="#000000" size="4">M</font><font face="Verdana" color="#000000" size="2">
      ediante este servicio usted puede conocer los últimos 10 movimientos de
      su cuenta. Esto puede resultarle muy útil si desea comprobar el
      resultado de una transferencia de fondos efectuada anteriormente. </font></strong></td>
    </tr>
  </table>
  </div>
</form>
<% Else
      ' Ejecuta los Ultimos Movimientos a partir de los datos seleccionados

  Cta_Debito=Request.Form("Cta_Debito")
  Cta_Credito=Request.Form("Cta_Credito")
  Importe=Request.Form("Imp")
  Observ=Request.Form("describe")
 
   X_MENSENV="0200"
   X_BBITMAP="F23880018A0080000000000006000000"           
   X_PAN=SPACE(15) & whois 
   X_PRCODE="390000"   
   X_AMOUNTRA="000000000000" 
   X_TDT=RIGHT("00" & CSTR(month(date)), 2)+RIGHT("00" & CSTR(day(date)), 2)+RIGHT("00" & CSTR(hour(time)), 2)+RIGHT("00" & CSTR(minute(time)), 2)+RIGHT("00" & CSTR(second(time)), 2)
   COUNTER=Application("NumVisits") 
   X_SYSTRACE=RIGHT("000000" & counter,6)
   X_TIMELOCA=MID(X_TDT,5,6)
   X_DATELOCA=MID(X_TDT,1,4)
   X_DATECAPT=MID(X_TDT,1,4)
   X_ACINCODE="  000000001"
   X_FOINCODE="  000000001"
   X_NUM_UNIQ=CSTR(YEAR(date)) & RIGHT("00" & CSTR(month(date)), 2) & RIGHT("00" & CSTR(day(date)), 2) & X_SYSTRACE
   X_RETRENUM=MID(X_NUM_UNIQ, 3, 12)
   X_RESPCODE="00"
   X_CURRCODE="840"  
   X_ACCOIDE1="28" & Request.Form("Cta_Debito") & "000000001  001" 
   X_ACCOIDE2="28" & SPACE(28)

   query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
   set rs = conn.Execute( query )

   X_DIRECCIO=rs("DIR_TDX25")   
   X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2
   
  query1 = "INSERT INTO M_MENSWI (ID_MENSAJE, PAN, TDT, IDCARCOD, IDCARTER, SYSTRACE, DES_MENSCE, Actioncode, Filename, Respcode, bit48resp, DIR_TDX25, NUM_UNIQCE, ACINCODE, TIMELOCA, DATELOCA, DATECAPT, RETRENUM, ENVIADO, CODIGO, FEC_ACCION, PROCESADO) VALUES ( '" & X_MENSENV & "', '" & X_PAN & "', '" & X_TDT & "', '', '', '" & X_SYSTRACE & "', '" & X_MENSAJE & "', '', '', '', '', '" & X_DIRECCIO & "', '" & X_NUM_UNIQ & "', '" & X_ACINCODE & "', '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_DATECAPT & "', '" & X_RETRENUM & "', '0', '39', {//}, '1' )"
  conn.Execute( query1 )%>
<IFRAME ALIGN=MIDDLE HEIGHT=390 allowtransparency="True" WIDTH=550 FRAMEBORDER=0 NORESIZE scrolling="no" SRC="10_Mov_Procesando.asp?count=37&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>&Cuenta=<%=Request.Form("Cta_Debito")%>"></iframe>
<%End If%>
</body>
</html>