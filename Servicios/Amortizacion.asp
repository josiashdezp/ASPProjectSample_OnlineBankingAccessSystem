<%  Response.Expires=0 
    Response.Buffer = True  
  %>  <html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Page-Exit" content="revealTrans(Duration=1,Transition=23)">
<title>Amortización de Préstamo</title>
<style type="text/css">
<!-- 
.style1 {
	font-family: Verdana;
	font-size: 12px;
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
<!-- #Include file = "Informa.asp"-->
<body style="background-color: transparent" leftmargin="0" topmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style11">Amortización de Fondos.</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
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
  
  ' Enviar mensaje 200 a la sucursal para que me envie el listado de adeudos

  tony = "Listado de Adeudos. Usuario: " & whois
  ltony = len(tony)
  X_MENSENV="0200"
  X_BBITMAP="F23880018A0180000000000006000000"           
  X_PAN=SPACE(15) & whois 
  X_PRCODE="510200"   
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
  F_ELECTR = ""
  ltony=ltony+14 
  X_OBSEV= Right("000" & ltony,3) & tony & "Comprobante:00"
  X_ACCOIDE1="28CUP15121" & whois & "00000000001  001" 
  X_ACCOIDE2="28CUP15121" & whois & "00000000001  001" 

  query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
  set rs = conn.Execute( query )

  X_DIRECCIO=rs("DIR_TDX25")   
  X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_OBSEV & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2

     query1 = "INSERT INTO M_MENSWI (ID_MENSAJE, PAN, TDT, IDCARCOD, IDCARTER, SYSTRACE, DES_MENSCE, Actioncode, Filename, Respcode, bit48resp, DIR_TDX25, NUM_UNIQCE, ACINCODE, TIMELOCA, DATELOCA, DATECAPT, RETRENUM, ENVIADO, CODIGO, FEC_ACCION, Procesado) VALUES ( '" & X_MENSENV & "', '" & X_PAN & "', '" & X_TDT & "', '', '', '" & X_SYSTRACE & "', '', '', '', '', '', '" & X_DIRECCIO & "', '" & X_NUM_UNIQ & "', '" & X_ACINCODE & "', '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_DATECAPT & "', '" & X_RETRENUM & "', '0', '02', {//}, '1')"
     conn.Execute query1 
     ' Ahora lo que hacemos es actualizar ese campo
         Set rsMessage = Server.CreateObject( "ADODB.RecordSet" )
		 
		 rsMessage.CursorType 		= 2 'adOpenKeyset
		 rsMessage.CursorLocation 	= 2 'adUseServer
		 rsMessage.Locktype 		= 3 'adLockDynamic  	
		 rsMessage.ActiveConnection = conn
		 rsMessage.Open "m_menswi"
		 rsMessage.Filter = "SYSTRACE = '" & X_SYSTRACE & "'"	 
         rsMessage("DES_MENSCE")= X_MENSAJE
		 rsMessage.UPDATE 
%>
<IFRAME ALIGN=MIDDLE  HEIGHT=460 WIDTH=100% FRAMEBORDER=0 NORESIZE SCROLLING=NO allowtransparency="True" SRC="Amort_Procesando_Adeudos.asp?count=37&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>"></iframe>
</body>
</html>