<% If  Session("UsrId") = "" then
      Response.Redirect("../Home.htm")
   else  
    Response.Expires=0 
    Response.Buffer = True  
   End If  %>  <html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>Solicitud de Chequeras</title>
<script language="JavaScript">
function solo_num(theComp)
{
  var checkOK = "0123456789";
  var checkStr = theComp.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  
  if (!allValid)
  {
    alert("Por favor escriba solo dígitos en este campo.");
    theComp.value = allNum
    theComp.focus();
    return (false);
  }
 return (true);
}

function FrontPage_Form1_Validator(theForm)
{
  if (theForm.Cnt.value == "")
  {
    alert("Especifique la cantidad de chequeras que desea imprimir.");
    theForm.Cnt.focus();
    return (false);
  }

  var checkOK = "0123456789";
  var checkStr = theForm.Cnt.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    
  }
  if (!allValid)
  {
    alert("Por favor, escriba solo dígitos en el campo \"Cantidad\".");
    theForm.Cnt.focus();
    return (false);
  }
return (true);
}
</script>
<style type="text/css">
<!-- 

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

.style16 {
	font-family: Verdana;
	font-size: 14px;
	color: #000000;
}
</style>
</head>
<body style="background-color: transparent" topmargin="0">
<!-- #Include file = "Informa.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" Align="Center">
  <tr>
    <td width="100%"><span class="style11">Solicitud de Chequeras.</span></td>
  </tr>
  <tr>
    <td width="100%"><hr size="1" color="#FFFFCC" align="left" width="98%">
    </td>
</table>
<% 
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4)
   
   Cuemay = "CUE" & sucu
   Application.Lock 
   Application("NumVisits") = Application("NumVisits") + 1 
   Application.Unlock 

  Set Conn = Server.CreateObject("ADODB.Connection")
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
  
  On Error Resume Next
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

  query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM " & Cuemay & " WHERE (Cod_Contra = '" & whois & "') AND (Cue_Inact <> .T.) AND (Cue_Cierre<> .T.) AND (Cue_bloq <> .T.) AND (CUE_SUBCUE <> '3360')"
  set rs = conn.Execute( query )
%>
<form METHOD="POST" ACTION="Chequera_Solicit.asp" onsubmit="return FrontPage_Form1_Validator(this)"name="FrontPage_Form1">
  <input type="hidden" name="hname" value="hvalue">
      <table border="0" width="100%" cellspacing="10" cellpadding="5" align="center">
        <%empresa=rs("Nom_client")%> 
      
        <tr> 
          <td align="left" valign="top" bgcolor="#FFD6D6" colspan="2"><font face="Verdana"><small><strong><small>&nbsp;&nbsp; 
                    <span class="style16">Cuenta a Debitar:</span></small></strong></small> 
            <select name="Cta_Debito" size="1" tabindex="1">
<% do while not rs.EOF
 if (mid(rs("cue_subcue"),1,3) <> "141") and (mid(rs("cue_subcue"),1,3) <> "151") then

     query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rs("sig_moneda") & "')"
     set rs2 = conn.Execute( query2 ) 
	 money = rs2("Cod_Moneda") 
	 cta = money & sucu & mid(rs("cue_subcue"), 3, 1) & whois & rs("des_cuenta")
%>     <option value="<%=cta%><%=Resto(cta)%>"><%=cta%><%=Resto(cta)%></option>
<% end if 
 rs.MoveNext 
  loop %> 
           </select></font></td>
        </tr>
        <tr align="center"> 
          <td height="30" colspan="2" align="center" valign="top" bgcolor="#FCE8AB"><span class="style16"><b><%=empresa%></b></span></td>
        </tr>
        <tr align="center"> 
          <td align="left" valign="top" bgcolor="#FFD6D6" colspan="2"><font face="Verdana" color="#0000FF"><strong><small><small class="style16">Tipo 
            de Chequera: </small><!--webbot bot="Validation" S-Display-Name="Cuenta a Acreditar" S-Data-Type="Integer" S-Number-Separators="x" B-Value-Required="TRUE" I-Minimum-Length="14" I-Maximum-Length="14" --> 
            <select name="TipCheq" size="1">
              <option value="001">Nominativo</option>
              <option value="101">Nominativo a la Orden</option>
              <option value="003">Certificado</option>
              <option value="103">Certificado a la Orden</option>
              <option value="010">Voucher Nominativo</option>
            </select>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            &nbsp; <small class="style16">Cantidad:</small>
            <input SIZE="2" NAME="Cnt" tabindex="3" MAXLENGTH="2" onkeyup="solo_num(Cnt)">
          </strong></font></td>
        </tr>
         <tr align="center"> 
          <td align="right" valign="top" colspan="2">
            <table border="0" width="100%" cellspacing="5" cellpadding="5">
              <tr> 
                <td align="center">
                  <input TYPE="submit" VALUE="Solicitar" tabindex="5" style="background-color: #9D2C4A; color: rgb(255,255,255); font-weight: bold" name="Button">
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input TYPE="reset" VALUE="Limpiar" tabindex="5" style="background-color: #9D2C4A; color: rgb(255,255,255); font-weight: bold">
                </td>
              </tr>
            </table>
          </td>
        </tr>
  </table>
  </center></div>
</form>
<% Else
      ' Ejecuta la Solicitud a partir de los datos seleccionados
 If Request.Form("Cta_Debito").Count = 0 Then %>
<table width="100%" align="center" cellpadding="5" bgcolor="#FCE8AB">
   <tr>
    <td align="left"><span class="style12">Error.  </span>
  </tr>
  <tr>
    <td align="center"><hr size="1">  
  </tr>
  <tr>
    <td><p align="center"><blink><font face="Verdana">
    <font color="#FF0000">
    <b>Usted no ha seleccionado ninguna Cuenta!!!.</b>
    </font> </font></blink></p>
    </td>
  </tr>
    <tr>
    <td align="center" class="style14"><a href="Chequera_Solicit.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
<% Else 
       queryCE = "SELECT ACT_CE FROM M_ESTSIS"
       set rsCE = conn.Execute( queryCE )
	 If rsce("ACT_CE") Then 
 
         Cant = Request.Form("Cnt") * 100
         X_MENSENV="0200"
         X_BBITMAP="F23880018A0180000000000006000000"           
         X_PAN=SPACE(15) & whois 
         X_PRCODE="510800"   
         X_AMOUNTRA=RIGHT("000000000000" & Cant,12) 
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
         X_CURRCODE=Request.Form("TipCheq")
         
         ELECTR = HEX (whois) & HEX (X_TDT) & HEX (Request.Form("Cnt") * 100)
         set AspObj = Server.CreateObject("Comp.CompCrypt")
         F_ELECTR = AspObj.Firmar( ELECTR )          
		 F_ELECTR = mid(F_ELECTR, 1, 20)
         
         lenfe=len(F_ELECTR)+14
         X_OBSEV= Right("000" & lenfe,3) & "Comprobante:" & Right("00" & lenfe,2) & F_ELECTR
         X_ACCOIDE1="28" & Request.Form("Cta_Debito") & "000000001  001" 
         X_ACCOIDE2="28" & "00000000000000000000001  001" 

      query = "SELECT DIR_TDX25 FROM C_BANCOS WHERE (COD_BANCO = '" & sucu & "')"
      set rs = conn.Execute( query )

      X_DIRECCIO=rs("DIR_TDX25")   
      X_MENSAJE=X_MENSENV & X_BBITMAP & "16" & RIGHT(X_PAN,16) & X_PRCODE & X_AMOUNTRA & X_TDT & X_SYSTRACE & X_TIMELOCA & X_DATELOCA & X_DATECAPT & "09" & mid(X_ACINCODE, 3, 12) & "09" & mid(X_FOINCODE, 3, 12) & X_RETRENUM & X_RESPCODE & X_OBSEV & X_CURRCODE & X_ACCOIDE1 & X_ACCOIDE2
         
	  query1 = "INSERT INTO M_MENSWI (ENVIADO, SYSTRACE, ID_MENSAJE, PAN, ACINCODE, TDT, IDCARCOD, IDCARTER, TIMELOCA, DATELOCA, RETRENUM, DES_MENSCE, ACTIONCODE, FILENAME, RESPCODE, BIT48RESP, NUM_UNIQCE, DIR_TDX25, CODIGO, DATECAPT, FEC_ACCION, PROCESADO) VALUES ('0', '" & X_SYSTRACE & "', '" & X_MENSENV & "', '" & X_PAN & "', '" & X_ACINCODE & "', '" & X_TDT & "', SPACE(15), SPACE(8), '" & X_TIMELOCA & "', '" & X_DATELOCA & "', '" & X_RETRENUM & "', '', SPACE(1), SPACE(17), SPACE(2), SPACE(0), '" & X_NUM_UNIQ & "',  '" & X_DIRECCIO & "', '08', '" & X_DATECAPT & "', {},'1')"
      conn.Execute( query1 )
	  
	' Ahora lo que hacemos es actualizar ese campo
         Set rsMessage = Server.CreateObject( "ADODB.RecordSet" )
		 
		 rsMessage.CursorType 		= 2 'adOpenKeyset
		 rsMessage.CursorLocation 	= 2 'adUseServer
		 rsMessage.Locktype 		= 3 'adLockDynamic  	
		 rsMessage.ActiveConnection = conn
		 rsMessage.Open "m_menswi"
		 rsMessage.Filter = "SYSTRACE = '" & X_SYSTRACE & "'"	 
         rsMessage("DES_MENSCE")= X_MENSAJE
		 rsMessage.UPDATE 	  %> 
      
      <IFRAME ALIGN=MIDDLE FRAMEBORDER=0 HEIGHT=460 WIDTH=100% NORESIZE SCROLLING=NONE allowtransparency="True" SRC="Chequera_Proces.asp?count=37&Systrace=<%=X_SYSTRACE%>&TDT=<%=X_TDT%>&Cta_debito=<%=Request.Form("Cta_Debito")%>&Cnt=<%=Request.Form("Cnt")%>&TipCheq=<%=Request.Form("TipCheq")%>"></iframe>
<%Else%> 
<table width="100%" cellpadding="5" cellspacing="5" bgcolor="#FCE8AB">
  <tr>
    <td align="left"><span class="style12">Error.  </span>
  </tr>
  <tr>
    <td align="center"><hr size="1">  
  </tr>
  <tr>
    <td align="center"><span class="style13">No existe
    conexión con el banco!!!</span></td>
  </tr>`
    <tr>
    <td align="center" class="style14"><a href="Chequera_Solicit.asp"><strong> < < Regresar >></strong></a>
  </tr>
</table>
	<%End If
 End If 
End If %> 
</center></div>
</body>
</html>