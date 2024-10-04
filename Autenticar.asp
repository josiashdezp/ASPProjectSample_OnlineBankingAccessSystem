<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/ConnString_BandecOnline.asp" -->
<%
User = Request("user")
Password = Request("password")
response.write User & "--" & Password
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_ConnString_BandecOnline_STRING
Recordset1.Source = "SELECT * FROM dbo.Administradores"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_ConnString_BandecOnline_STRING
Recordset2.Source = "SELECT *  FROM dbo.Administradores"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset2.EOF)) 
%>
<table width="100%%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><%=(Recordset2.Fields.Item("Nombre").Value)%></td>
    <td><%=(Recordset2.Fields.Item("Apellidos").Value)%></td>
    <td><%=(Recordset2.Fields.Item("Login").Value)%></td>
    <td><%=(Recordset2.Fields.Item("Contraseña").Value)%></td>
    <td><%=(Recordset2.Fields.Item("Correo").Value)%></td>
  </tr>
</table>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset2.MoveNext()
Wend
%>

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
