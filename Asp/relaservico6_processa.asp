<%@ Language=VBScript%>
<% Server.ScriptTimeOut = 3600 %>
<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

IF strLoginRede <> "EDAR" and strLoginRede <> "PRSSILV" and strLoginRede <> "FRPORTO"  THEN
	msg = "<p align=center><b><font color=#000080 face=Arial Black size=6>Sistema NewCLA</font></b></p>"
	msg = msg & "<p align=center><b><font color=#000080 face=Arial Black size=4>Em Testes</font></b></p>"
	Response.write msg
	response.end
END IF

'Option Explicit
'Response.Expires = 0
'Área de definição de variáveis
Dim str_sql, tipo, FldRS
Dim cont
Dim ObjConn

'Espaço para código ASP (Funções etc)
str_sql = Request.Form("str_sql")
tipo = Request.Form("radio1")
%>

<html>
<head>
	<title>Processa</title>
</head>

<body>
<% 


Set ObjConn = Server.CreateObject ("ADODB.Connection")
'Abre a conecção
	
'Conecta com NEWCLA
ObjConn.ConnectionString = "file name=d:\inetpub\ConexaoSQL\NewCLA.udl"
ObjConn.Open

IF tipo = "Tipo2" then
	Set ObjRS = CreateObject("ADODB.RecordSet")
	ObjRS.ActiveConnection = ObjConn
	ObjRS.Open(str_sql)%>
	 <table bordercolor="#0033ff" border="1">
	  <tr>
	   <td>Nº</td>
	   <% For Each FldRS In ObjRS.Fields %>
	    <td><%= Response.Write(FldRS.Name) %></td>
	   <% Next %>
	  </tr>
		<%cont = 1
		  While Not ObjRS.EOF %>
	      <tr>
	       <td><font color=red><%=cont%></font></td>
		   <% For Each FldRS In ObjRS.Fields %>
		    <td><%= Response.Write(FldRS.Value) %></td>
		   <% Next %>
		  </tr>
		<%cont = cont + 1
		  ObjRS.MoveNext
		WEnd
		ObjRS.Close
		ObjConn.Close
		Set ObjRS = Nothing
		Set ObjConn = Nothing %>	  
	 </table>
<% else 
	'Bloqueio de segurança:
	if instr(str_sql,"update") = 1 then 
		if instr(str_sql,"where") = 0 then
			response.write "ERRO. Tenha mais atenção!"
			response.end
		end if
	end if
	
	if instr(str_sql,"delete") = 1 then
		if instr(str_sql,"where") = 0 then
			response.write "ERRO. Tenha mais atenção!"
			response.end
		end if
	end if
	
	if instr(str_sql,"truncate") = 1 then 'Update
			response.write "Não permitido"
			response.end
	end if
	
	if instr(str_sql,"drop") = 1 then 'Update
			response.write "Não permitido"
			response.end
	end if
	
%>


<%ObjConn.Execute(CStr(str_sql))%>
SQL executado na base:<br>
<%=str_sql %>
<%ObjConn.close
set ObjConn = Nothing
END IF %>


</body>
</html>