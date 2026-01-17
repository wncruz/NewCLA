<!--#include file="../inc/adovbs.inc"-->

<%
server.ScriptTimeout = 28800
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
If strLoginRede <> "IMPLEME" and strLoginRede <> "JOAOFNS" then
	Response.Write "USUÁRIO INVÁLIDO!!!"
	Response.end
end if

Dim db, RS, str, fld

	Set db = server.createobject("ADODB.Connection")
	db.ConnectionString = "file name=d:\inetpub\ConexaoSQL\NewCLA.udl"
	db.CursorLocation = adUseClient
	db.open

if Trim(Request.Form("t")) <> "" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	Set RS = db.Execute(request("t"))

	str = "<table width=415><tr><td align=right>" & RS.RecordCount & " rows</td></tr></table>"

	str = str & "<TABLE cellspacing=0 rules=all bordercolorlight=#ffffff bordercolordark=#003399 width=500>" & vbCrLf

	str = str & "<TR>" & vbCrLf
	For Each fld In Rs.Fields
		str = str & "<TH>&nbsp;" & fld.Name & "</TH>" & vbCrLf
	Next
	str = str & "</TR>" & vbCrLf

	cor = "#dddddd"
	Do Until RS.EOF

		if cor = "#dddddd" then
			cor = "#eeeeee"
		else
			cor = "#dddddd"
		end if
		
		str = str & "<TR>" & vbCrLf
		For Each fld In Rs.Fields
			str = str & "<TD bgcolor=" & cor & ">&nbsp;" & fld.Value & "</TD>" & vbCrLf
		Next
		str = str & "</TR>" & vbCrLf

		RS.MoveNext
	Loop

	str = str & "</TABLE>" & vbCrLf
	
	str = str & "<table width=415><tr><td align=right>" & RS.RecordCount & " rows</td></tr></table>"

	db.Close
end if
%>
<tr><td align="center" class="titulo" width="780">Usuario: <%=strLoginRede%> -Consulta<br><br></td></tr>

<form method="post">
<tr>
	<td align="center" width="780"><textarea name="t" cols="108" rows="8" style="font-family : courier new; font-weight : normal;"><%=request("t")%></textarea></td>
</tr>
<tr>
	<td align="center" width="780"><br><input type="submit" class="button" value="Enviar"></td>
</tr>

<tr>
<td><br><%=str%></td>
</tr>

</BODY>
</HTML>
