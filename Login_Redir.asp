<%@ Language=VBScript %>
<%  Whois = Request("login")

	  Set Conn = Server.CreateObject("ADODB.Connection")
	  conn.Open "File Name="&Server.MapPath("Connections/ConnString_BandecOnline.udl")
 
	 query1="Select Cambiar_Password, Fecha_Validacion, Emp_Nombre from tbl_contratos Where (login='" & Whois & "')"
	 set rs1 = conn.Execute( query1 )
 
	 if not rs1.eof then
	   Mi_Fecha = rs1("Fecha_Validacion")
	   Empresa = rs1("Emp_Nombre")
	   
	   If not rs1("Cambiar_Password")=True AND DateDiff("d", Mi_Fecha, Date()) < 90 then
	 	Session("UsrId") = Whois
		Session("CheckConnectTime")=0
	    Response.write "<script language="&CHR(34)&"JavaScript1.2"&CHR(34)&">"&CHR(13)&CHR(10)
    	Response.write "parent.location.href='servicios.asp?Empresa="& Empresa & "'" & CHR(13) & CHR(10)
        Response.write  "</script>" & CHR(13) & CHR(10)
       Else
        Response.Redirect("change.asp?Expire=True")
       End IF  'end del if not rsl
     Else
	   Session.Abandon()
	   Response.Redirect("Login_Error.asp?Code=06")
	 End If 'end del if not rsl.eof
%>