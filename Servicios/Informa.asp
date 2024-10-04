<% 
 whois = Session("UsrId")
  if not whois="" then
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open "File Name=" & Server.MapPath("../Connections/ConnString_BandecOnline.udl")
  query1 = "Select Cambiar_Password, Fecha_Validacion from tbl_Contratos Where (login='" & Whois & "')"
  
  
    set rs1 = Conn.Execute( query1 )
    Mi_Fecha = rs1("Fecha_Validacion")
    If not rs1("Cambiar_Password")=True AND DateDiff("d", Mi_Fecha, Date()) < 90 then
      date_conn = Date()
	  time_conn = Time()
	  
      position = inStr( Request.ServerVariables("URL"),"/")
	  service=mid(Request.ServerVariables("URL"),position+11)
  
	  query = "INSERT INTO Informa (Tipo_Autenticacion, Usuario, Dir_Remota, Fecha_Conexion, Hora_Conexion, Servicio) VALUES ('" & Request.ServerVariables("AUTH_TYPE") & "',  '" & Session("UsrId") & "', '" & Request.ServerVariables("Remote_Addr") & "', " & "convert(smalldatetime,'" & Date() & "',101)" & ", '" & Time() & "', '" & service & "')"
	  
	  Conn.Execute( query )
	 Else
       Response.Clear
       Response.Redirect("change.asp?Expire=True")
     End IF  
	 else 'No existe session, lo redirecciono pal login
      Response.write "<script language="&CHR(34)&"JavaScript1.2"&CHR(34)&">"&CHR(13)&CHR(10)
      Response.write "parent.location.href='../barcode.asp?Msg=2'"&CHR(13)&CHR(10)
      Response.write "</script>"&CHR(13)&CHR(10)
	 end if
%>