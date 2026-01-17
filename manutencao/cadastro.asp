
<%@ Language=VBScript%>
<% Server.ScriptTimeOut = 3600 %>
<%

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
	<title>Atualização da base - CLA Produção</title>
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
<% else %>

<%ObjConn.Execute(CStr(str_sql))%>
SQL executado na base:<br>
<%=str_sql %>
<%ObjConn.close
set ObjConn = Nothing
END IF %>


</body>
</html>