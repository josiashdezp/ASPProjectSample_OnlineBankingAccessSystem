<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<%
	Const DTSSQLStgFlag_Default = 0
	Const DTSStepExecResult_Failure = 1
	
	Dim oPkg, oStep, sMessage, bStatus
	
	Set oPkg = Server.CreateObject("DTS.Package")
	oPkg.LoadFromSQLServer "bdc-cf","BandecOnline","bandeconline",DTSSQLStgFlag_Default,"","",""," NOMBRE DEL PAQUETE "
	oPkg.Execute()
	
	bStatus = True
	
	For Each oStep In oPkg.Steps
		sMessage = sMessage & "<p> Step [" & oStep.Name & "] "
		If oStep.ExecutionResult = DTSStepExecResult_Failure Then
			sMessage = sMessage & " failed<br>"
			bStatus = False
		Else
			sMessage = sMessage & " succeeded<br>"
		End If
		sMessage = sMessage & "Task """ & oPkg.Tasks(oStep.TaskName).Description & """</p>"
	Next
	
	If bStatus Then
		sMessage = sMessage & "<p>Package [" & oPkg.Name & "] succeeded</p>"
	Else
		sMessage = sMessage & "<p>Package [" & oPkg.Name & "] failed</p>"
	End If
	
	Response.Write sMessage
	Response.Write "<p>Done</p>"
%>