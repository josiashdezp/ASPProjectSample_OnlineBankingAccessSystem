<%  Response.Expires=0 
    Response.Buffer = True  
   %>  <html>
   <!--#include file="Informa.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Chequeo de Conexión con Sucursal.</title>
</head>
<body style="background-color: transparent"  leftmargin="0">
<% Serv = Request("Serv") 
   if Session("CheckConnectTime")<>0 then
	 if datediff("n",Session("CheckConnectTime"),Time()) < 10   then
   Response.Write "Servicio" & Serv   
	    Select Case Serv
		      Case "TR"   Response.Redirect("Transferencias_Fondos.asp") '0 es Transferencias de Fondos
		      Case "AP"   Response.Redirect("Aporte.asp") '1 es Aporte al presupuesto
			  case "CH"   Response.Redirect("Chequera_Solicit.asp") '3 es solicitar chequera
       End Select
	 end if
   end if 
   
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4) 
   Application.Lock 
   Application("NumVisits") = Application("NumVisits") + 1 
   Application.Unlock 
 set conn = Server.CreateObject( "ADODB.Connection" )
   conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl")

   X_MENSENV="0200"
   X_BBITMAP="F23880018A0080000000000006000000"           
   X_PAN=SPACE(15) & whois 
   X_PRCODE="500000"   
   X_AMOUNTRA="000000000000" 
   X_TDT=RIGHT("00" & CSTR(month(date)), 2)+RIGHT("00" & CSTR(day(date)), 2)+RIGHT("00" & CSTR(hour(time)), 2)+RIGHT("00" & CSTR(minute(time)), 2)+RIGHT("00" & CSTR(second(time)), 2)
   COUNTER=Application("NumVisits") 
   X_SYSTRACE=RIGHT("000000" & Application("NumVisits"),6)
   X_TIMELOCA=MID(X_TDT,5,6)
   X_DATELOCA=MID(X_TDT,1,4)
   X_DATECAPT=MID(X_TDT,1,4)
   X_ACINCODE="  000000001"
   X_FOINCODE="  000000001"
   X_NUM_UNIQ=CSTR(YEAR(date)) & RIGHT("00" & CSTR(month(date)), 2) & RIGHT("00" & CSTR(day(date)), 2) & X_SYSTRACE
   X_RETRENUM=MID(X_NUM_UNIQ, 3, 12)
   X_RESPCODE="00"
   X_CURRCODE="840"  
   X_ACCOIDE1="28" & SPACE(28)
   X_ACCOIDE2="28" & SPACE(28)

   query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
   set rs = conn.Execute( query )

   X_DIRECCIO=rs("DIR_TDX25")   
   X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2
   
   query1 = "INSERT INTO M_MENSWI (ID_MENSAJE, PAN, TDT, IDCARCOD, IDCARTER, SYSTRACE, DES_MENSCE, Actioncode, Filename, Respcode, bit48resp, DIR_TDX25, NUM_UNIQCE, ACINCODE, TIMELOCA, DATELOCA, DATECAPT, RETRENUM, ENVIADO, CODIGO, FEC_ACCION, PROCESADO) VALUES ( '" & X_MENSENV & "', '" & X_PAN & "', '" & X_TDT & "', '', '', '" & X_SYSTRACE & "', '" & X_MENSAJE & "', '', '', '', '', '" & X_DIRECCIO & "', '" & X_NUM_UNIQ & "', '" & X_ACINCODE & "', '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_DATECAPT & "', '" & X_RETRENUM & "', '0', '39', {//}, '1' )"
  set rs1 = conn.Execute( query1 )
  %>
<IFRAME frameborder="0" allowtransparency="true" height="100%" width="100%" scrolling="auto" SRC="Consultar_Sucursal_Procesando.asp?count=25&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>&Sucu=<%=Sucu%>&Serv=<%=Serv%>"></IFRAME>
</body>
</html>