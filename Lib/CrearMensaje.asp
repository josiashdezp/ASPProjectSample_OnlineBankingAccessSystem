<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%'Datos generales que se utilizan para calcular no importa si es una o todas las cuentas 

  set conn = Server.CreateObject( "ADODB.Connection" )
  conn.Open "File Name="&Server.MapPath("../Connections/Conn_VFProx.udl") 
  
   the_date 	= Request.Querystring("the_date")   
   the_date1 	= Request.Querystring("the_date1")
   email 		= Request.Form("email")   
   cc 			=  Request.Form("cc")
   NroCuenta	= Request.Querystring("cta")
   
   sucu = mid(Session("UsrId"), 1 , 4)
   whois = mid(Session("UsrId"), 5 , 4) 

	histor = "HIST" & sucu 
	Cuemay = "CUE" & sucu

   
   '**************************************************************************************************
'VAMOS A BUSCAR LOS NÚMEROS DE CUENTAS, ESTOS E GUARDAN EN UN ARREGLO.

   dim Cuentas 			  	'Este es el listado de las cuentas que vamos a buscar acá se guarda 
	   Cuentas = Array("","","","","","","","","","")
	   contador = 0	 		'inicializamos el contador de la cantidad de cuentas que se va a buscar el estado de cuentas

if (NroCuenta = "Todas") then 'ahora vamos a ver si se va a buscar todas la cuentas o solamente una.
	
     
	   	Function Resto(cuenta)
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
	    
      
	  'PARA BUSCAR LAS CUENTAS QUE VAMOS A REVISAR, PRIMERO VEREMOS SI TIENE 
	  'QUE BUSCAR CUENTAS O SUB - CUENTAS:

	  if Request("SubCuentas")<> true then	  'BUSCAR LAS CUENTAS.
		file = "CUE" & sucu 
	 	query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM "& file &" WHERE (Cod_Contra = '" & whois & "') AND (Cue_Inact = .F.) AND (Cue_Cierre = .F.)"
 	 	set rsCuentas = conn.Execute( query ) 
	 
	 	  do while not rsCuentas.EOF        
		  if (rsCuentas("cue_subcue") = "3360") or (rsCuentas("cue_subcue") = "3210") or (rsCuentas("cue_subcue") = "3280") or (rsCuentas("cue_subcue") = "3290") then 
	    	if (rsCuentas("cue_subcue") = "3360") then 
			   contador=contador+1	
			   localcta = rsCuentas("sig_moneda") & rsCuentas("cue_subcue") & rsCuentas("tip_contra") & whois & rsCuentas("des_cuenta")    
    	       Cuentas(contador) = localcta 
	   	    else
			  contador=contador+1	
			  query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rsCuentas("sig_moneda") & "')"
        	  set rsCuentas2 = conn.Execute( query2 ) 
			  money = rsCuentas2("Cod_Moneda") 
 			  localcta = money & sucu & mid(rsCuentas("cue_subcue"), 3, 1) & whois & rsCuentas("des_cuenta") 
			  Cuentas(contador) = localcta & Resto(localcta)
		    end if 
	      end if 
		  rsCuentas.MoveNext 
		  loop
		  rsCuentas.Close
		  rsCuentas2.Close       
		  ' HASTA AQUÍ EL DE BUSCAR LAS CUENTAS
	  else  
	  'DE LO CONTRARIO VAMOS A BUSCAR LA SUBCUENTAS
	     query = "SELECT * FROM subclien WHERE (CUE_SUCUR = '" & sucu & "') and (COD_CONTRA = '" & whois & "')" 
  		 set subc = conn.Execute(query) 
	     do while not subc.EOF     
	     file = "CUE" & subc("sucu") 
    	 query = "SELECT Sig_Moneda, Cue_Subcue, Tip_Contra, Des_Cuenta, Nom_Client FROM "& file &" WHERE (Cod_Contra = '" & subc("whois") & "') AND (Cue_Inact = .F.) AND (Cue_Cierre = .F.) AND (Cue_Cierre = .F.)"
	     set rsCuentas = conn.Execute( query ) 
	      do while not rsCuentasCuentas.EOF        
		      if (rsCuentas("cue_subcue") = "3360") or (rsCuentas("cue_subcue") = "3210") or (rsCuentas("cue_subcue") = "3280") or (rsCuentas("cue_subcue") = "3290") then 
			      if (rsCuentas("cue_subcue") = "3360") then 
    		     	 contador=contador+1
					 localcta = rsCuentas("sig_moneda") & rsCuentas("cue_subcue") & rsCuentas("tip_contra") & subc("whois") & rsCuentas("des_cuenta")    
					 Cuentas(contador) = localcta
				  else   
			         contador=contador+1	
					 query2 = "SELECT Cod_Moneda FROM c_moneda WHERE (SIG_MONEDA = '" & rsCuentas("sig_moneda") & "')"
        			 set rsCuentas2 = conn.Execute( query2 ) 
	  			  	 money = rsCuentas2("Cod_Moneda") 
					 localcta = money & subc("sucu") & mid(rsCuentas("cue_subcue"), 3, 1) & subc("whois") & rsCuentas("des_cuenta")    	
					 Cuentas(contador) = localcta & Resto(localcta)
    	         end if 
			  end if    
		 rsCuentas.MoveNext 
		 loop
		 subc.MoveNext
	     loop
		 rsCuentas.Close
		 rsCuentas2.Close
		 'ACÁ TERMINA EL DE LAS SUBCUENTAS
	  end if 'END DEL IF REQUEST SUBCUENTAS <> TRUE
	  
	  
else  ' si es solo una cuenta hacemos esto
   Cuentas(1) = NroCuenta
   contador=contador+1	
end if

'********************************************************
'Ahora vamos a armar el estado de cuentas y crear el mensaje.

   	mensaje="<html><head>"
    mensaje=mensaje & " <title>Estado de Cuentas por Email</title> "&vbcrlf
    mensaje=mensaje & " </head><body><div align=center> "&vbcrlf
	mensaje=mensaje & " <P><B><div align=left><font color=#000000 face=Verdana size=3>Estado de Cuentas.</font></b></p>"&vbcrlf
	mensaje=mensaje & " <hr size=1 color=#000000 align=left> "&vbcrlf
	
	for i = 1 to contador

			 Cta = Cuentas(i)

			queryb= "SELECT PLAZA FROM C_BANCOS WHERE (COD_BANCO= '" & SUCU & "')"
		    set rsv = conn.Execute( queryb )
			
			'Este es el encabezado donde dice el número de la cuenta y la fecha
			mensaje=mensaje & " <p align=right><b><font color=#000000 face=Verdana size=2>BANDEC, SUCURSAL "&sucu&", "&rsv("PLAZA")&"</font><br>"&vbcrlf
			mensaje=mensaje & " <font color=#000000 face=Verdana size=2>Fecha Emisi&oacute;n: " & Date() & " </font><br>"&vbcrlf
			mensaje=mensaje & " <font color=#000000 face=Verdana size=2>Cuenta: " & cta & " </font></b></p>"&vbcrlf
			mensaje=mensaje & " <div align=center><center><table border=0 cellpadding=0 cellspacing=5 width='100%'><tr><td align=center colspan=6><hr color=#000000 size=1></td></tr>"&vbcrlf
			mensaje=mensaje & " <tr><td width=8% align=center><p align=center><font face=Arial size='2'><strong>Fecha Contable</strong></span></font></td>"&vbcrlf
		    mensaje=mensaje & " <td width=10% align=center><p align=center><font face=Arial size='2'1px><strong>Referencia Corriente</strong></span></font></td>"&vbcrlf
			mensaje=mensaje & " <td width=10% align=center><p align=center><font face=Arial size='2'><strong>Referencia Original</strong></span></font></td>"&vbcrlf
			mensaje=mensaje & " <td width=50% align=center><font face=Arial size='2'><strong>Observaciones&nbsp;&nbsp;</strong></span></font></td><td width=10 align=center></td><td width=19 align=center><font face=Arial size='2'><span style=text-transform: uppercase><strong>Movimientos</strong></span></font></td></tr>"&vbcrlf
			mensaje=mensaje & " <tr><td colspan=6><hr color=#000000 size=1></td></tr><tr><td width=12></td><td width=16></td><td width=15></td><td colspan=2><p align=right><strong><span style=text-transform: uppercase><font face=Arial size='2'>Saldo Anterior: </font></span></strong></td>"&vbcrlf

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

	'************************
	
	  querysa = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & Des_cuenta & "') AND (FEC_CONTAB >= {^" & the_date & "}) AND (COD_ASIENT='120') Order by Fec_Contab"	
  set rssa = conn.Execute(querysa)

   If CDbl(rssa("IMP_ASIENT")) < 0 then cod = "Cr" else cod = "Db" End If
   
	Imp_asient = FormatNumber(Abs(CDbl(rssa("IMP_ASIENT"))),2) & cod
	mensaje=mensaje & " <td align=right><font face=Arial><font size=2>" & Imp_asient & " </font></font></td></tr> "&vbcrlf
	mensaje=mensaje & " <tr><td width=12></td><td width=16></td><td width=15></td><td width=28><font size='2'></font></td><td colspan=2><font size='2'><hr color=#000000 width=100 size=1 align=right></font></td></tr></table></center></div>"&vbcrlf
	
	query1 = "SELECT FEC_CONTAB, REF_CORRIE, REF_ORIGIN, OBSERV, IMP_ASIENT, COD_ASIENT FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB>={^" & the_date & "}) AND (FEC_CONTAB<={^" & the_date1 & "}) AND ((IsNull(COD_ASIENT))  OR ((not IsNull(COD_ASIENT)) AND (COD_ASIENT <> '120') AND (COD_ASIENT <> '121') AND (COD_ASIENT <> '122') AND (COD_ASIENT <> '123') AND (COD_ASIENT <> '125') AND (COD_ASIENT <> '126'))) Order by Fec_Contab ASC"	
    set rs1 = conn.Execute(query1)  
	
 Do While not rs1.eof 
    
	If rs1("COD_ASIENT") <> "124" OR IsNull(rs1("COD_ASIENT")) then
    mensaje=mensaje & " <div align=center> "&vbcrlf 

	mensaje=mensaje & " <div align=center><center> "&vbcrlf
    mensaje=mensaje & " <table border=0 cellpadding=0 cellspacing=5 width='100%'><tr> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><p align=center><font face=Arial size='2'> " & rs1("FEC_CONTAB")& " </font></td> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><font face=Arial><font size='2'> " & rs1("REF_CORRIE")& " </font></font></td> "&vbcrlf
	mensaje=mensaje & " <td width=10% align=center valign=Top><p align=center><font face=Arial size='2'> " & rs1("REF_ORIGIN") & " </font></td> "&vbcrlf
	 Obs = rs1("observ")  
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
	 End If 
	 mensaje=mensaje & " <td width=50% align=center><p align=right><font face=Arial size='2'> " &  Obs & " </font></td><td width=3></td></center> "&vbcrlf
  	 If CDbl(rs1("IMP_ASIENT")) < 0 then cod = "Cr" else cod = "Db" End If
	 Imp_asient = FormatNumber(abs(CDbl(rs1("IMP_ASIENT"))),2) & cod 
     mensaje=mensaje & " <td align=right valign=Top><font size=2><font face=Arial>" & Imp_Asient & "&nbsp;&nbsp;</font></font></td></tr></table></div></div> "&vbcrlf
	Else
	 mensaje=mensaje & " <div align=center><center><table border=0 cellpadding=0 cellspacing=5 width='100%'><tr> "&vbcrlf
	 mensaje=mensaje & " <td width=10% align=center valign=Top><small><small><font face=Arial> " & rs1("FEC_CONTAB") & " </font></small></small></td> "&vbcrlf
	 mensaje=mensaje & " <td align=left><strong><small><small><font face=Arial> No hubo movimientos en esta fecha. </font></small></small></strong></td></tr></table></center></div> "&vbcrlf
	End If
	
rs1.MoveNext 
Loop 
	mensaje=mensaje & " <div align=center> "&vbcrlf

	queryfecha = "SELECT FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB >= {^" & The_date & "}) AND (FEC_CONTAB <= {^" & The_date1 & "})AND (COD_ASIENT='121')"
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
	mensaje=mensaje & " <table border=0 cellpadding=0 cellspacing=0 width='100%'><tr><td width=9 rowspan=2></td><td width=9 rowspan=2></td><td width=9 rowspan=2></td><td width=9 rowspan=2></td><td colspan=2 rowspan=2><font size='2'></font></td><td rowspan=2 width=18><font size='2'><hr color=#000000 size=1 width=100 align=right></font></td></tr> "&vbcrlf
	mensaje=mensaje & " <tr></tr><tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size='2'><span style=text-transform: uppercase><b>Saldo final:</b></span></font></td>"&vbcrlf
	
	queryscont = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {^" & The_Date1 & "}) AND (COD_ASIENT='121') Order By Fec_Contab DESC"
    set rsscont = conn.Execute( queryscont )
    
	If CDbl(rsscont("IMP_ASIENT")) < 0 then cod = "Cr" else cod = "Db" End If
	
	Imp_Asient = FormatNumber(Abs(CDbl(rsscont("IMP_ASIENT"))),2) & cod
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	

   if (subcue = "3280") or (subcue = "3290") then Saldo = "Fondo aprobado:" else  Saldo = "Sobregiro autorizado:" end if
   
	mensaje=mensaje & " <tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size='2'><span style=text-transform: uppercase><b>" & Saldo & "</b></span></font></td>"&vbcrlf

    querysf = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {^" & The_Date1 & "'}) AND (COD_ASIENT='126') Order By Fec_Contab DESC"
    set rssf = conn.Execute( querysf )
	
	
	If rssf.Eof then Impt=0 else Impt=rssf("IMP_ASIENT") end if
	
	If IMPT < 0 then cod = "Cr" else cod = "Db" End If
	
    Imp_Asient = FormatNumber(Abs(IMPT),2) & cod
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	
	mensaje=mensaje & " <tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size='2'><span style=text-transform: uppercase><b>Fondo Reservado: </b></span></font></td>"&vbcrlf
	
	querysdisp = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {^" & The_Date1 & "}) AND (COD_ASIENT='125') Order By Fec_Contab DESC"
    set rssdisp = conn.Execute( querysdisp )
    

	If CDbl(rssdisp("IMP_ASIENT")) < 0 then cod = "Cr" else cod = "Db" End If
	
	Imp_Asient = FormatNumber(Abs(CDbl(rssdisp("IMP_ASIENT"))),2) & cod
	
	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	mensaje=mensaje & " <tr><td width=9></td><td width=9></td><td width=9></td><td width=9></td><td align=right colspan=2><font face=Arial size='2'><span style=text-transform: uppercase><b>Fondo Disponible:</b></span></font></td>"&vbcrlf


	
	querysdisp = "SELECT IMP_ASIENT, FEC_CONTAB FROM " & histor & " WHERE (SIG_MONEDA='" & money & "') AND (CUE_SUBCUE='" & subcue & "') AND (COD_CONTRA='" & whois & "') AND (DES_CUENTA='" & des_cuenta & "') AND (FEC_CONTAB <= {^" & The_Date1 & "}) AND (COD_ASIENT='123') Order By Fec_Contab DESC"
	set rssdisp = conn.Execute( querysdisp )

	If CDbl(rssdisp("IMP_ASIENT")) < 0 then cod = "Cr" else cod = "Db" End If
	Imp_Asient = FormatNumber(Abs(CDbl(rssdisp("IMP_ASIENT"))),2) & cod

	mensaje=mensaje & " <td align=right><font size=2><font face=Arial> " & Imp_Asient & "&nbsp;&nbsp;&nbsp;&nbsp;</font></font></td></tr> "&vbcrlf		
	mensaje=mensaje & " </table></center></div></div> "&vbcrlf
next
	
'************************************************************
'EL PIE DE PÁGINA ES ESTE
	mensaje=mensaje & " <hr size=1 color=#000000 align=left width='100%'>"&vbcrlf
	mensaje=mensaje & " <font face=verdana size='2'><div align=center><b>Servicio de Banca Electr&oacute;nica. Bandec Online.</b></div></font> "&vbcrlf

'*************************************************************

'Y AHORA MANDARLO POR EL CORREO
 dim MyFileSystem, MyTextFile, MyFile
	
	set MyFileSystem =  Server.CreateObject("Scripting.FileSystemObject")
		
		MyFileName   =  Server.MapPath("../") & "\Var\MSG"& whois &".html"

		if MyFileSystem.FileExists(MyFileName)=true	then
			set MyTextFile = MyFileSystem.OpenTextFile(MyFileName,2,true,-2)		
		else
			set MyTextFile = MyFileSystem.CreateTextFile( MyFileName )		 
		End if
		
    	MyTextFile.WriteLine(mensaje)
		MyTextFile.close
		
	set MyFileSystem = nothing
	set MyTextFile   = nothing
	
	 Response.Redirect("EnviarMensaje.aspx?MSG=" & whois & "&To=" & email & "&Cc=" & cc)	  %>