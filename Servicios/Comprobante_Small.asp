<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>BANDEC ONLINE. Comprobante de Transferencia</title>

</head>
<style media="all" type="text/css" >

 .Comprobante 
 {  font-family:Arial, Helvetica, sans-serif;
 	font-size:11pt; 
 } 	
 
 .Titulo 
 {
  font-size:10pt;
  text-align:center; 
  }
  
 .Nomb_Comp
  {
  text-decoration:underline;
  font-size:9pt;
  }
  
</style>
<%

'Los tres tipos de comprobantes que se imprimen por esta vía son:
'      
'	    TRANS (Comprobante de Trnasferencias)
'		APORT (Comprobante de Aportes)
'		CHEQ  (Comprobante de Solicitus de Chequeras)
'		AMORT  (Compribante de Amortización)
	 	

	Tipo_Comp = Request("Tipo")
	Cta_Debito = Request("Cta_Debito")
	Resultado = Request("Resultado")	

   dim Comprobante 			 
	   Comprobante = Array("","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","")
	   

  if (Tipo_Comp = "TRANSF") then   'comprobante de transferencia

   		FC = Request("FC")
		RC = Request("RC")
		Importe = Request("Importe")		
		Cta_Credito = Request("Cta_Credito")		
		Describe = Request("Describe")
		ELECTR = Request("ELECTR")
		MAX = 20

	   Comprobante(1) = "<div class='Titulo'> BANCO DE CR&Eacute;DITO Y COMERCIO"
	   Comprobante(2) = "BANDEC ONLINE"
	   Comprobante(4) = "<span class='Nomb_Comp'> COMPROBANTE DE TRANSFERENCIA </span></div>"
   	   Comprobante(5) = "Fecha de Solicitud:  " & FC & " "
   	   Comprobante(6) = "Cuenta Db: " & Cta_Debito & ""
   	   Comprobante(7) = "Referencia Corriente: " & RC & ""
   	   Comprobante(8) = "Cuenta Cr: " & Cta_Credito & ""
	   Comprobante(9)  = "Importe: $" & FormatNumber(Importe,2) & ""
	   Comprobante(10)  = "Observaciones: "
	   Comprobante(11)  = describe & ""
	   Comprobante(12)  = "Resultado de la Transferencia: "
   	   Comprobante(13)  = resultado & ""
	   Comprobante(14)  = "Comprobante de la Transferencia: "	   
	   Comprobante(15)  = ELECTR & ""
	   Comprobante(16)  = "  HECHO:"
	   Comprobante(17)  = ""
	   Comprobante(18)  = "  AUTORIZADO:"
	   Comprobante(19)  = ""
	   Comprobante(20)  = "  Para cualquier reclamaci&oacute;n presente este comprobante"
	   
 else
   if (Tipo_Comp = "APORT") then   'comprobante de aportes
 
		Importe = Request("Importe")       	
		Parrafo = Request("Parrafo")
		NIT = Request("NIT")
		Nom_Client = Request("Nom_Client")
		Pcpal = Request("Pcpal")
		Recargo = Request("Recargo")
		Multa = Request("Multa")
		PD = Request("PD")
		PH = Request("PH")
		HD = Request("HD")
		TP = Request("TP")
		RF= Request("RF")																														
		II = Request("II")																														
		TI = Request("TI")																														
		IO = Request("IO")																														
		SUC = Request("SUC")																														
		Firma = Request("Firma")
		FC = Request("FC")
		RC = Request("RC")	
		MAX = 37																																												

 	   Comprobante(1) = "<div class='Titulo'> BANCO DE CR&Eacute;DITO Y COMERCIO"
	   Comprobante(2) = "BANDEC ONLINE"
	   Comprobante(3) = "<span class='Nomb_Comp'> COMPROBANTE DE APORTE </span> </div>"
  	   Comprobante(4) = ""
   	   Comprobante(5) = "Nro de Ident. Tribut.: " & NIT
   	   
       If mid(Parrafo,7,1)=2 then Nro = "- 03" else Nro = "- 04" End IF
	   
	   Comprobante(6) = "CR " & Nro  
   	   Comprobante(7) = "Deb&iacute;tese a: " & Nom_Client
   	   Comprobante(8) = "C&oacute;digo de la cuenta: "
	   Comprobante(9)  = Cta_Debito
	   Comprobante(10)  = "Tributo: " & mid(Parrafo,8)
	   Comprobante(11)  = "Código: "  & mid(PARRAFO,1,7)
	   Comprobante(12)  = "Importe: $" & FormatNumber(Importe,2)
   	   Comprobante(13)  = "Principal: $" & Pcpal 
	   Comprobante(14)  = "Recargo: " & Recargo
	   Comprobante(15)  = "Multa o Sanción: " & Multa
	   Comprobante(16)  = "Per&iacute;odo a liquidar:(DD/MM/AAAA)"
	   Comprobante(17)  = "Desde: " & mid(PD,1,2) & "/" & mid(PD,4,2) & "/" & mid(PD,7,4)
	   Comprobante(18)  = "Hasta: " & mid(PH,1,2) & "/" & mid(PH,4,2) & "/" & mid(PH,7,4)
	   Comprobante(19)  = "Referencia de Pago:"
	   Comprobante(20)  = "TP: " & TP & " D: " & RF
	   Comprobante(21)  = "Base Imp.:" & "$ " & II
	   Comprobante(22)  = "TI: $ " & TI
	   Comprobante(23)  = "Importe de la Obligación:"
	   Comprobante(24)  = "$ " & IO
	   Comprobante(25)  = "Resultado:"
	   Comprobante(26)  = Resultado	   	   	   	   
	   Comprobante(27)  = "Tramitado en la sucursal: " & SUC 
	   Comprobante(28)  = "Referencia: " & RC
	   Comprobante(29)  = "Fecha Cont.: " & FC
	   Comprobante(30)  = "Comprobante: " 
	   Comprobante(31)  = Firma   	   
   	   Comprobante(32)  = " "    	   
	   Comprobante(33)  = "HECHO:"	   	   	   	   
	   Comprobante(34)  = ""	   	   	   	   
	   Comprobante(35)  = "AUTORIZADO:"	   	   	   	   
	   Comprobante(36)  = ""	   	   	   	   
	   Comprobante(37)  = "Para cualquier reclamaci&oacute;n presente este comprobante"  	
	ELSE
	   if (Tipo_Comp = "CHEQ") then    'COMPROBANTE DE CHEQUERAS

	    TipCheq = Request("TipCheq")
		Cnt = Request("Cnt")
		ELECTR = Request("ELECTR")
		MAX = 18

	   Comprobante(1) = "<div class='Titulo'> BANCO DE CR&Eacute;DITO Y COMERCIO"
	   Comprobante(2) = "BANDEC ONLINE"
	   Comprobante(3) = "<span class='Nomb_Comp'> COMPROBANTE DE SOLICITUD DE CHEQUERA </span> </div>"
	   Comprobante(4) = ""
   	   Comprobante(5) = "Fecha de Solicitud: " & Date
   	   Comprobante(6) = "Cuenta Db: " & Cta_Debito 
   	   Comprobante(7) = "Tipo de Cheque: " 
   	   Comprobante(8) = "   " & TipCheq
	   Comprobante(9)  = "Cantidad: " & FormatNumber(Cnt,0)
	   Comprobante(10)  = "Resultado de la Solicitud:"
   	   Comprobante(11)  = resultado
	   Comprobante(12)  = "Comprobante de la Solicitud:"	   
	   Comprobante(13)  = ELECTR 
	   Comprobante(14)  = "  HECHO:"
	   Comprobante(15)  = ""
	   Comprobante(16)  = "  AUTORIZADO:"
	   Comprobante(17)  = ""
	   Comprobante(18)  = "  Para cualquier reclamaci&oacute;n presente este comprobante"
	  else
	     if (Tipo_Comp = "AMORT") then 
		 
		 FC = Request("FC")
		 RC = Request("RC")	
		 Cta_Credito = Request("Cta_Credito")
		 Firma = Request("Firma")	
		 Principal = Request("Principal")		 		 	 
		 Interes = Request("Interes")
		 MAX = 19
		    
			
	   Comprobante(1) = "<div class='Titulo'> BANCO DE CR&Eacute;DITO Y COMERCIO"
	   Comprobante(2) = "BANDEC ONLINE"
	   Comprobante(3) = "<span class='Nomb_Comp'> COMPROBANTE DE AMORTIZACI&Oacute;N </span> </div>"
   	   Comprobante(4) = ""
   	   Comprobante(5) = "Fecha Contable: " & FC
   	   Comprobante(6) = "Cuenta Db: " & Cta_Debito 
   	   Comprobante(7) = "Ref. Corriente: " 
   	   Comprobante(8) = "Cuenta Cr: " & Cta_Credito 
	   Comprobante(9)  = "Princ. Amort.: $ " & Principal
	   Comprobante(10)  = "Interes Amort.: $ " & Interes
   	   Comprobante(11)  = "Resultado: "
   	   Comprobante(12)  = resultado
	   Comprobante(13)  = "Comprobante de la Transferencia:"	   
	   Comprobante(14)  = Firma 
	   Comprobante(15)  = "  HECHO:"
	   Comprobante(16)  = ""
	   Comprobante(17)  = "  AUTORIZADO:"
	   Comprobante(18)  = ""
	   Comprobante(19)  = "  Para cualquier reclamaci&oacute;n presente este comprobante"
		 end if
	  END IF
   end if
end if 
 

  TEXTO = ""

  For I = 1 To MAX
   TEXTO = TEXTO & "<br>" & Comprobante(I)
  Next
 
 %>
<body onload="window.print();" class="Comprobante">
<!-- window.close(); -->
<%
 For I = 1 To MAX
  Response.Write Comprobante(I)
  Response.Write "<br>"
Next   %>  
</body>
</html>